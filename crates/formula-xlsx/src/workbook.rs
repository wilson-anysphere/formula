use std::collections::HashMap;

use formula_model::charts::{Chart, ChartKind, ChartModel, ChartType};
use roxmltree::Document;

use crate::charts::parse_chart;
use crate::drawingml::charts::{
    extract_chart_object_refs, parse_chart_color_style, parse_chart_ex, parse_chart_space,
    parse_chart_style, ChartColorStyleParseError, ChartDiagnostic, ChartDiagnosticSeverity,
    ChartExParseError, ChartObject, ChartParts, ChartSpaceParseError, ChartStyleParseError,
    OpcPart,
};
use crate::drawingml::extract_chart_refs;
use crate::package::XlsxPackage;
use crate::path::{rels_for_part, resolve_target};
use crate::relationships::parse_relationships;

#[derive(Debug, thiserror::Error)]
pub enum ChartExtractionError {
    #[error("missing part: {0}")]
    MissingPart(String),
    #[error("part is not valid UTF-8: {0}: {1}")]
    XmlNonUtf8(String, #[source] std::str::Utf8Error),
    #[error("failed to parse XML: {0}: {1}")]
    XmlParse(String, #[source] roxmltree::Error),
    #[error("invalid XML structure: {0}")]
    XmlStructure(String),
}

impl XlsxPackage {
    pub fn extract_charts(&self) -> Result<Vec<Chart>, ChartExtractionError> {
        let drawing_to_sheet = self.drawing_to_sheet_map().unwrap_or_default();
        let mut charts = Vec::new();

        for drawing_part in self
            .part_names()
            .filter(|name| name.starts_with("xl/drawings/drawing") && name.ends_with(".xml"))
            .map(str::to_string)
            .collect::<Vec<_>>()
        {
            let sheet_info = drawing_to_sheet.get(&drawing_part);
            let sheet_name = sheet_info.map(|info| info.name.clone());
            let sheet_part = sheet_info.and_then(|info| info.part.clone());

            let drawing_xml = match self.part(&drawing_part) {
                Some(bytes) => bytes,
                None => continue,
            };

            let drawing_refs = extract_chart_refs(drawing_xml, &drawing_part)?;
            if drawing_refs.is_empty() {
                continue;
            }

            let drawing_rels_part = rels_for_part(&drawing_part);
            let drawing_rels = self.part(&drawing_rels_part);
            let drawing_rel_map = match drawing_rels {
                Some(xml) => parse_relationships(xml, &drawing_rels_part)?
                    .into_iter()
                    .map(|r| (r.id.clone(), r))
                    .collect::<HashMap<_, _>>(),
                None => HashMap::new(),
            };

            for drawing_ref in drawing_refs {
                let relationship = drawing_rel_map.get(&drawing_ref.rel_id);
                let chart_part = relationship
                    .map(|r| resolve_target(&drawing_part, &r.target))
                    .filter(|target| self.part(target).is_some());

                let parsed = match chart_part
                    .as_deref()
                    .and_then(|chart_part| self.part(chart_part).map(|bytes| (chart_part, bytes)))
                {
                    Some((chart_part, bytes)) => parse_chart(bytes, chart_part).unwrap_or(None),
                    None => None,
                };

                let (chart_type, title, series) = match parsed {
                    Some(parsed) => (parsed.chart_type, parsed.title, parsed.series),
                    None => (
                        ChartType::Unknown {
                            name: "unparsed".to_string(),
                        },
                        None,
                        Vec::new(),
                    ),
                };

                charts.push(Chart {
                    sheet_name: sheet_name.clone(),
                    sheet_part: sheet_part.clone(),
                    drawing_part: drawing_part.clone(),
                    chart_part: chart_part.map(|s| s.to_string()),
                    rel_id: drawing_ref.rel_id,
                    chart_type,
                    title,
                    series,
                    anchor: drawing_ref.anchor,
                });
            }
        }

        Ok(charts)
    }

    pub fn extract_chart_objects(&self) -> Result<Vec<ChartObject>, ChartExtractionError> {
        let drawing_to_sheet = self.drawing_to_sheet_map().unwrap_or_default();
        let mut chart_objects = Vec::new();

        for drawing_part in self
            .part_names()
            .filter(|name| name.starts_with("xl/drawings/drawing") && name.ends_with(".xml"))
            .map(str::to_string)
            .collect::<Vec<_>>()
        {
            let sheet_info = drawing_to_sheet.get(&drawing_part);
            let sheet_name = sheet_info.map(|info| info.name.clone());
            let sheet_part = sheet_info.and_then(|info| info.part.clone());

            let drawing_xml = match self.part(&drawing_part) {
                Some(bytes) => bytes,
                None => continue,
            };

            let drawing_chart_refs = extract_chart_object_refs(drawing_xml, &drawing_part)?;
            if drawing_chart_refs.is_empty() {
                continue;
            }

            let drawing_rels_part = rels_for_part(&drawing_part);
            let drawing_rel_map = match self.part(&drawing_rels_part) {
                Some(xml) => parse_relationships(xml, &drawing_rels_part)?
                    .into_iter()
                    .map(|r| (r.id.clone(), r))
                    .collect::<HashMap<_, _>>(),
                None => HashMap::new(),
            };

            for drawing_ref in drawing_chart_refs {
                let mut diagnostics = Vec::new();

                let chart_part_path = match drawing_rel_map.get(&drawing_ref.rel_id) {
                    Some(rel) => resolve_target(&drawing_part, &rel.target),
                    None => {
                        diagnostics.push(ChartDiagnostic {
                            severity: ChartDiagnosticSeverity::Error,
                            message: format!(
                                "missing drawing relationship for chart reference {}",
                                drawing_ref.rel_id
                            ),
                            part: Some(drawing_rels_part.clone()),
                            xpath: None,
                        });
                        String::new()
                    }
                };

                let chart_bytes = if chart_part_path.is_empty() {
                    Vec::new()
                } else {
                    match self.part(&chart_part_path) {
                        Some(bytes) => bytes.to_vec(),
                        None => {
                            diagnostics.push(ChartDiagnostic {
                                severity: ChartDiagnosticSeverity::Error,
                                message: format!("missing chart part: {chart_part_path}"),
                                part: Some(chart_part_path.clone()),
                                xpath: None,
                            });
                            Vec::new()
                        }
                    }
                };

                let (chart_rels_path, chart_rels_bytes) = if chart_part_path.is_empty() {
                    (None, None)
                } else {
                    let rels = rels_for_part(&chart_part_path);
                    match self.part(&rels) {
                        Some(bytes) => (Some(rels), Some(bytes.to_vec())),
                        None => (None, None),
                    }
                };

                let mut chart_ex_part = None;
                let mut style_part = None;
                let mut colors_part = None;
                let mut user_shapes_part = None;

                if let (Some(chart_rels_path), Some(chart_rels_bytes)) =
                    (chart_rels_path.as_deref(), chart_rels_bytes.as_deref())
                {
                    let rels = parse_relationships(chart_rels_bytes, chart_rels_path)?;
                    for rel in rels {
                        let target_path = resolve_target(&chart_part_path, &rel.target);
                        let rel_type = rel.type_.to_ascii_lowercase();
                        let rel_target = rel.target.to_ascii_lowercase();

                        if chart_ex_part.is_none()
                            && (rel_type.contains("chartex") || rel_target.contains("chartex"))
                        {
                            chart_ex_part = Some(target_path);
                            continue;
                        }

                        if style_part.is_none() && is_chart_style_relationship(&rel_type, &rel_target)
                        {
                            style_part = Some(target_path);
                            continue;
                        }

                        if colors_part.is_none()
                            && is_chart_colors_relationship(&rel_type, &rel_target)
                        {
                            colors_part = Some(target_path);
                            continue;
                        }

                        if user_shapes_part.is_none()
                            && is_chart_user_shapes_relationship(&rel_type, &rel_target)
                        {
                            user_shapes_part = Some(target_path);
                            continue;
                        }
                    }
                }

                let chart = OpcPart {
                    path: chart_part_path.clone(),
                    rels_path: chart_rels_path.clone(),
                    rels_bytes: chart_rels_bytes,
                    bytes: chart_bytes,
                };

                let chart_ex = chart_ex_part.and_then(|path| match self.part(&path) {
                    Some(bytes) => {
                        let rels_path = rels_for_part(&path);
                        let rels_bytes = self.part(&rels_path).map(|bytes| bytes.to_vec());
                        Some(OpcPart {
                            path,
                            rels_path: rels_bytes.as_ref().map(|_| rels_path),
                            rels_bytes,
                            bytes: bytes.to_vec(),
                        })
                    }
                    None => {
                        diagnostics.push(ChartDiagnostic {
                            severity: ChartDiagnosticSeverity::Warning,
                            message: format!("missing chartEx part: {path}"),
                            part: Some(path),
                            xpath: None,
                        });
                        None
                    }
                });

                let style = style_part.and_then(|path| match self.part(&path) {
                    Some(bytes) => Some(OpcPart {
                        path,
                        rels_path: None,
                        rels_bytes: None,
                        bytes: bytes.to_vec(),
                    }),
                    None => {
                        diagnostics.push(ChartDiagnostic {
                            severity: ChartDiagnosticSeverity::Warning,
                            message: format!("missing chart style part: {path}"),
                            part: Some(path),
                            xpath: None,
                        });
                        None
                    }
                });

                let colors = colors_part.and_then(|path| match self.part(&path) {
                    Some(bytes) => Some(OpcPart {
                        path,
                        rels_path: None,
                        rels_bytes: None,
                        bytes: bytes.to_vec(),
                    }),
                    None => {
                        diagnostics.push(ChartDiagnostic {
                            severity: ChartDiagnosticSeverity::Warning,
                            message: format!("missing chart colors part: {path}"),
                            part: Some(path),
                            xpath: None,
                        });
                        None
                    }
                });

                let user_shapes = user_shapes_part.and_then(|path| match self.part(&path) {
                    Some(bytes) => {
                        let rels_path = rels_for_part(&path);
                        let rels_bytes = self.part(&rels_path).map(|bytes| bytes.to_vec());
                        Some(OpcPart {
                            path,
                            rels_path: rels_bytes.as_ref().map(|_| rels_path),
                            rels_bytes,
                            bytes: bytes.to_vec(),
                        })
                    }
                    None => {
                        diagnostics.push(ChartDiagnostic {
                            severity: ChartDiagnosticSeverity::Warning,
                            message: format!("missing chart userShapes part: {path}"),
                            part: Some(path),
                            xpath: None,
                        });
                        None
                    }
                });

                let chart_space_model = if chart.path.is_empty() || chart.bytes.is_empty() {
                    None
                } else {
                    match parse_chart_space(&chart.bytes, &chart.path) {
                        Ok(model) => Some(model),
                        Err(err) => {
                            diagnostics.push(ChartDiagnostic {
                                severity: ChartDiagnosticSeverity::Warning,
                                message: format!(
                                    "failed to parse chartSpace part {}: {}",
                                    chart.path,
                                    format_chart_space_error(&err)
                                ),
                                part: Some(chart.path.clone()),
                                xpath: None,
                            });
                            None
                        }
                    }
                };

                let chart_ex_model = if let Some(chart_ex_part) = chart_ex.as_ref() {
                    match parse_chart_ex(&chart_ex_part.bytes, &chart_ex_part.path) {
                        Ok(model) => {
                            if let ChartKind::Unknown { name } = &model.chart_kind {
                                if let Some(kind) = name.strip_prefix("ChartEx:") {
                                    if kind.trim().is_empty() || kind == "unknown" {
                                        diagnostics.push(ChartDiagnostic {
                                            severity: ChartDiagnosticSeverity::Warning,
                                            message: "chartEx part detected but chart kind could not be inferred"
                                                .to_string(),
                                            part: Some(chart_ex_part.path.clone()),
                                            xpath: None,
                                        });
                                    }
                                }
                            }
                            Some(model)
                        }
                        Err(err) => {
                            diagnostics.push(ChartDiagnostic {
                                severity: ChartDiagnosticSeverity::Warning,
                                message: format!(
                                    "failed to parse ChartEx part {}: {}",
                                    chart_ex_part.path,
                                    format_chart_ex_error(&err)
                                ),
                                part: Some(chart_ex_part.path.clone()),
                                xpath: None,
                            });
                            None
                        }
                    }
                } else {
                    None
                };

                let mut model = match (chart_space_model, chart_ex_model) {
                    (Some(chart_space_model), Some(chart_ex_model)) => Some(merge_chart_models(
                        chart_space_model,
                        chart_ex_model,
                        &chart.path,
                        chart_ex
                            .as_ref()
                            .map(|p| p.path.as_str())
                            .unwrap_or_default(),
                        &mut diagnostics,
                    )),
                    (Some(chart_space_model), None) => Some(chart_space_model),
                    (None, Some(chart_ex_model)) => Some(chart_ex_model),
                    (None, None) => None,
                };

                if let Some(model) = model.as_mut() {
                    if let Some(style_part) = style.as_ref() {
                        match parse_chart_style(&style_part.bytes, &style_part.path) {
                            Ok(style_model) => model.style_part = Some(style_model),
                            Err(err) => diagnostics.push(ChartDiagnostic {
                                severity: ChartDiagnosticSeverity::Warning,
                                message: format!(
                                    "failed to parse chartStyle part {}: {}",
                                    style_part.path,
                                    format_chart_style_error(&err)
                                ),
                                part: Some(style_part.path.clone()),
                                xpath: None,
                            }),
                        }
                    }

                    if let Some(colors_part) = colors.as_ref() {
                        match parse_chart_color_style(&colors_part.bytes, &colors_part.path) {
                            Ok(colors_model) => model.colors_part = Some(colors_model),
                            Err(err) => diagnostics.push(ChartDiagnostic {
                                severity: ChartDiagnosticSeverity::Warning,
                                message: format!(
                                    "failed to parse chartColorStyle part {}: {}",
                                    colors_part.path,
                                    format_chart_color_style_error(&err)
                                ),
                                part: Some(colors_part.path.clone()),
                                xpath: None,
                            }),
                        }
                    }
                }

                chart_objects.push(ChartObject {
                    sheet_name: sheet_name.clone(),
                    sheet_part: sheet_part.clone(),
                    drawing_part: drawing_part.clone(),
                    anchor: drawing_ref.anchor,
                    drawing_frame_xml: drawing_ref.drawing_frame_xml,
                    parts: ChartParts {
                        chart,
                        chart_ex,
                        style,
                        colors,
                        user_shapes,
                    },
                    model,
                    diagnostics,
                });
            }
        }

        Ok(chart_objects)
    }

    fn drawing_to_sheet_map(&self) -> Result<HashMap<String, SheetInfo>, ChartExtractionError> {
        const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let workbook_part = "xl/workbook.xml";
        let workbook_xml = match self.part(workbook_part) {
            Some(xml) => xml,
            None => return Ok(HashMap::new()),
        };

        let workbook_xml = std::str::from_utf8(workbook_xml)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?;
        let workbook_doc = Document::parse(workbook_xml)
            .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;

        let workbook_rels_part = "xl/_rels/workbook.xml.rels";
        let workbook_rels_xml = match self.part(workbook_rels_part) {
            Some(xml) => xml,
            None => return Ok(HashMap::new()),
        };

        let workbook_rels = parse_relationships(workbook_rels_xml, workbook_rels_part)?;
        let mut workbook_rel_map = HashMap::new();
        for rel in workbook_rels {
            workbook_rel_map.insert(rel.id.clone(), rel);
        }

        let mut out = HashMap::new();

        for sheet_node in workbook_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "sheet")
        {
            let sheet_name = match sheet_node.attribute("name") {
                Some(name) => name.to_string(),
                None => continue,
            };
            let sheet_rid = match sheet_node
                .attribute((REL_NS, "id"))
                .or_else(|| sheet_node.attribute("r:id"))
                .or_else(|| sheet_node.attribute("id"))
            {
                Some(id) => id,
                None => continue,
            };

            let sheet_target = match workbook_rel_map.get(sheet_rid) {
                Some(rel) => resolve_target(workbook_part, &rel.target),
                None => continue,
            };

            let sheet_xml = match self.part(&sheet_target) {
                Some(xml) => xml,
                None => continue,
            };
            let sheet_xml = std::str::from_utf8(sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet_target.clone(), e))?;
            let sheet_doc = Document::parse(sheet_xml)
                .map_err(|e| ChartExtractionError::XmlParse(sheet_target.clone(), e))?;

            let drawing_rids: Vec<String> = sheet_doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "drawing")
                .filter_map(|n| {
                    n.attribute((REL_NS, "id"))
                        .or_else(|| n.attribute("r:id"))
                        .or_else(|| n.attribute("id"))
                })
                .map(|s| s.to_string())
                .collect();

            if drawing_rids.is_empty() {
                continue;
            }

            let sheet_rels_part = rels_for_part(&sheet_target);
            let sheet_rels_xml = match self.part(&sheet_rels_part) {
                Some(xml) => xml,
                None => continue,
            };
            let sheet_rels = parse_relationships(sheet_rels_xml, &sheet_rels_part)?;
            let rel_map: HashMap<_, _> =
                sheet_rels.into_iter().map(|r| (r.id.clone(), r)).collect();

            for drawing_rid in drawing_rids {
                let drawing_target = match rel_map.get(&drawing_rid) {
                    Some(rel) => resolve_target(&sheet_target, &rel.target),
                    None => continue,
                };

                out.insert(
                    drawing_target,
                    SheetInfo {
                        name: sheet_name.clone(),
                        part: Some(sheet_target.clone()),
                    },
                );
            }
        }

        Ok(out)
    }
}

#[derive(Debug, Clone)]
struct SheetInfo {
    name: String,
    part: Option<String>,
}

fn is_chart_style_relationship(rel_type: &str, rel_target: &str) -> bool {
    if rel_type.contains("chartstyle") {
        return true;
    }
    // Fallback to filename heuristic for producers that omit the relationship type.
    rel_target.ends_with(".xml")
        && rel_target.contains("style")
        && !rel_target.ends_with("styles.xml")
}

fn is_chart_colors_relationship(rel_type: &str, rel_target: &str) -> bool {
    if rel_type.contains("chartcolorstyle") {
        return true;
    }
    rel_target.ends_with(".xml") && rel_target.contains("colors")
}

fn is_chart_user_shapes_relationship(rel_type: &str, rel_target: &str) -> bool {
    if rel_type.contains("chartusershapes") {
        return true;
    }
    // Fallback to filename heuristic for producers that omit the relationship type.
    rel_target.ends_with(".xml") && rel_target.contains("drawing")
}

fn format_chart_space_error(err: &ChartSpaceParseError) -> String {
    err.to_string()
}

fn format_chart_ex_error(err: &ChartExParseError) -> String {
    err.to_string()
}

fn format_chart_style_error(err: &ChartStyleParseError) -> String {
    err.to_string()
}

fn format_chart_color_style_error(err: &ChartColorStyleParseError) -> String {
    err.to_string()
}

fn merge_chart_models(
    chart_space: ChartModel,
    chart_ex: ChartModel,
    chart_space_part: &str,
    chart_ex_part: &str,
    diagnostics: &mut Vec<ChartDiagnostic>,
) -> ChartModel {
    let chart_space_series_len = chart_space.series.len();
    let chart_ex_series_len = chart_ex.series.len();
    let chart_space_axes_len = chart_space.axes.len();
    let chart_ex_axes_len = chart_ex.axes.len();

    let mut merged = chart_space;

    // Chart kind: prefer ChartEx if it appears to be a meaningful subtype (e.g. "ChartEx:waterfall")
    // rather than the parser placeholder (e.g. "ChartEx:unknown").
    if chart_ex_kind_is_specific(&chart_ex.chart_kind) {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: format!(
                "model.chart_kind: using ChartEx {chart_ex_kind:?} (chartSpace was {chart_space_kind:?})",
                chart_ex_kind = &chart_ex.chart_kind,
                chart_space_kind = &merged.chart_kind,
            ),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
        merged.chart_kind = chart_ex.chart_kind.clone();
    } else {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: format!(
                "model.chart_kind: using chartSpace {chart_space_kind:?} (ChartEx was {chart_ex_kind:?})",
                chart_space_kind = &merged.chart_kind,
                chart_ex_kind = &chart_ex.chart_kind,
            ),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
    }

    // Title / legend: prefer ChartEx when present (future-proofing, as ChartEx parsing is still incomplete).
    if chart_ex.title.is_some() {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: "model.title: using ChartEx".to_string(),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
        merged.title = chart_ex.title.clone();
    } else {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: "model.title: using chartSpace".to_string(),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
    }

    if chart_ex.legend.is_some() {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: "model.legend: using ChartEx".to_string(),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
        merged.legend = chart_ex.legend.clone();
    } else {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: "model.legend: using chartSpace".to_string(),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
    }

    // Series / axes: fall back to chartSpace when ChartEx doesn't produce as many objects.
    if chart_space_series_len > chart_ex_series_len {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: format!(
                "model.series: using chartSpace (chartSpace={chart_space_series_len}, ChartEx={chart_ex_series_len})",
            ),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
    } else {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: format!(
                "model.series: using ChartEx (ChartEx={chart_ex_series_len}, chartSpace={chart_space_series_len})",
            ),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
        merged.series = chart_ex.series.clone();
    }

    if chart_space_axes_len > chart_ex_axes_len {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: format!(
                "model.axes: using chartSpace (chartSpace={chart_space_axes_len}, ChartEx={chart_ex_axes_len})",
            ),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
    } else {
        diagnostics.push(ChartDiagnostic {
            severity: ChartDiagnosticSeverity::Info,
            message: format!(
                "model.axes: using ChartEx (ChartEx={chart_ex_axes_len}, chartSpace={chart_space_axes_len})",
            ),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
        merged.axes = chart_ex.axes.clone();
    }

    // Preserve whichever diagnostics are available from both models.
    merged.diagnostics.extend(chart_ex.diagnostics);

    merged
}

fn chart_ex_kind_is_specific(kind: &ChartKind) -> bool {
    let ChartKind::Unknown { name } = kind else {
        // A fully-modeled ChartEx kind would be represented as a concrete enum variant. Treat it as specific.
        return true;
    };

    let Some(kind) = name.strip_prefix("ChartEx:") else {
        return false;
    };

    let kind = kind.trim();
    !kind.is_empty() && !kind.eq_ignore_ascii_case("unknown")
}
