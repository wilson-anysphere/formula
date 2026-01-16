use std::collections::HashMap;

use formula_model::charts::{
    Chart, ChartAnchor, ChartColorStylePartModel, ChartKind, ChartModel, ChartStylePartModel,
    ChartType, PlotAreaModel,
};
use formula_model::drawings::Anchor as DrawingAnchor;
use roxmltree::{Document, Node};

use crate::charts::legacy_parsed_chart_from_model;
use crate::drawingml::charts::{
    extract_chart_object_refs, parse_chart_color_style, parse_chart_ex, parse_chart_space,
    parse_chart_style, ChartColorStyleParseError, ChartDiagnostic, ChartDiagnosticLevel,
    ChartExParseError, ChartObject, ChartParts, ChartSpaceParseError, ChartStyleParseError,
    OpcPart,
};
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
        let chart_objects = self.extract_chart_objects()?;
        let mut charts = Vec::with_capacity(chart_objects.len());

        for chart_object in chart_objects {
            let parsed = chart_object
                .model
                .as_ref()
                .map(legacy_parsed_chart_from_model);

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

            let chart_part = if chart_object.parts.chart.path.is_empty() {
                None
            } else {
                Some(chart_object.parts.chart.path.clone())
            };
            let anchor = chart_anchor_from_drawing_anchor(chart_object.anchor);

            charts.push(Chart {
                sheet_name: chart_object.sheet_name,
                sheet_part: chart_object.sheet_part,
                drawing_part: chart_object.drawing_part,
                chart_part,
                rel_id: chart_object.drawing_rel_id,
                chart_type,
                title,
                series,
                anchor,
            });
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
            let mut external_drawing_rel_targets = HashMap::new();
            let drawing_rel_map = match self.part(&drawing_rels_part) {
                Some(xml) => {
                    let mut map = HashMap::new();
                    for rel in parse_relationships(xml, &drawing_rels_part)? {
                        if is_external_target_mode(rel.target_mode.as_deref()) {
                            external_drawing_rel_targets.insert(rel.id.clone(), rel.target.clone());
                            continue;
                        }
                        map.insert(rel.id.clone(), rel);
                    }
                    map
                }
                None => HashMap::new(),
            };

            for drawing_ref in drawing_chart_refs {
                let mut diagnostics = Vec::new();

                if let Some(external_target) = external_drawing_rel_targets.get(&drawing_ref.rel_id)
                {
                    diagnostics.push(ChartDiagnostic {
                        level: ChartDiagnosticLevel::Warning,
                        message: format!(
                            "chart reference {} points to external URI: {}",
                            drawing_ref.rel_id, external_target
                        ),
                        part: Some(drawing_rels_part.clone()),
                        xpath: None,
                    });

                    chart_objects.push(ChartObject {
                        sheet_name: sheet_name.clone(),
                        sheet_part: sheet_part.clone(),
                        drawing_part: drawing_part.clone(),
                        drawing_rel_id: drawing_ref.rel_id,
                        drawing_object_id: drawing_ref.drawing_object_id,
                        drawing_object_name: drawing_ref.drawing_object_name,
                        anchor: drawing_ref.anchor,
                        drawing_frame_xml: drawing_ref.drawing_frame_xml,
                        parts: ChartParts {
                            chart: OpcPart {
                                path: String::new(),
                                rels_path: None,
                                rels_bytes: None,
                                bytes: Vec::new(),
                            },
                            chart_ex: None,
                            style: None,
                            colors: None,
                            user_shapes: None,
                        },
                        model: None,
                        diagnostics,
                    });
                    continue;
                }

                let normalized_drawing_part = resolve_target(&drawing_part, "");

                let chart_part_path = match drawing_rel_map.get(&drawing_ref.rel_id) {
                    Some(rel) => {
                        let target = normalize_relationship_target(&rel.target);
                        let resolved = resolve_target(&drawing_part, &target);
                        if resolved == normalized_drawing_part {
                            diagnostics.push(ChartDiagnostic {
                                level: ChartDiagnosticLevel::Error,
                                message: format!(
                                    "invalid chart relationship target for {}: {}",
                                    drawing_ref.rel_id, rel.target
                                ),
                                part: Some(drawing_rels_part.clone()),
                                xpath: None,
                            });
                            String::new()
                        } else {
                            resolved
                        }
                    }
                    None => {
                        diagnostics.push(ChartDiagnostic {
                            level: ChartDiagnosticLevel::Error,
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
                                level: ChartDiagnosticLevel::Error,
                                message: format!("missing chart part: {chart_part_path}"),
                                part: Some(chart_part_path.clone()),
                                xpath: None,
                            });
                            Vec::new()
                        }
                    }
                };

                let (chart_rels_path, chart_rels_bytes) =
                    if chart_part_path.is_empty() || chart_bytes.is_empty() {
                        (None, None)
                    } else {
                        let rels_path = rels_for_part(&chart_part_path);
                        let rels_bytes = self.part(&rels_path).map(|bytes| bytes.to_vec());
                        if rels_bytes.is_none() {
                            diagnostics.push(ChartDiagnostic {
                                level: ChartDiagnosticLevel::Warning,
                                message: format!("missing chart relationships part: {rels_path}"),
                                part: Some(rels_path.clone()),
                                xpath: None,
                            });
                        }
                        (Some(rels_path), rels_bytes)
                    };

                let mut chart_ex_part = None;
                let mut style_part = None;
                let mut colors_part = None;
                let mut user_shapes_part = None;

                if let (Some(chart_rels_path), Some(chart_rels_bytes)) =
                    (chart_rels_path.as_deref(), chart_rels_bytes.as_deref())
                {
                    let rels = parse_relationships(chart_rels_bytes, chart_rels_path)?;
                    let normalized_chart_part = resolve_target(&chart_part_path, "");
                    for rel in rels {
                        if is_external_target_mode(rel.target_mode.as_deref()) {
                            continue;
                        }
                        let normalized_target = normalize_relationship_target(&rel.target);
                        let target_path = resolve_target(&chart_part_path, &normalized_target);
                        if target_path == normalized_chart_part {
                            diagnostics.push(ChartDiagnostic {
                                level: ChartDiagnosticLevel::Warning,
                                message: format!(
                                    "invalid chart relationship target for {}: {}",
                                    rel.id, rel.target
                                ),
                                part: Some(chart_rels_path.to_string()),
                                xpath: None,
                            });
                            continue;
                        }

                        if chart_ex_part.is_none()
                            && (crate::ascii::contains_ignore_case(&rel.type_, "chartex")
                                || crate::ascii::contains_ignore_case(&normalized_target, "chartex"))
                        {
                            chart_ex_part = Some(target_path);
                            continue;
                        }

                        if style_part.is_none()
                            && is_chart_style_relationship(&rel.type_, &normalized_target)
                        {
                            style_part = Some(target_path);
                            continue;
                        }

                        if colors_part.is_none()
                            && is_chart_colors_relationship(&rel.type_, &normalized_target)
                        {
                            colors_part = Some(target_path);
                            continue;
                        }

                        if user_shapes_part.is_none()
                            && is_chart_user_shapes_relationship(&rel.type_, &normalized_target)
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
                        if rels_bytes.is_none() {
                            diagnostics.push(ChartDiagnostic {
                                level: ChartDiagnosticLevel::Warning,
                                message: format!("missing chartEx relationships part: {rels_path}"),
                                part: Some(rels_path.clone()),
                                xpath: None,
                            });
                        }
                        Some(OpcPart {
                            path,
                            rels_path: Some(rels_path),
                            rels_bytes,
                            bytes: bytes.to_vec(),
                        })
                    }
                    None => {
                        diagnostics.push(ChartDiagnostic {
                            level: ChartDiagnosticLevel::Warning,
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
                            level: ChartDiagnosticLevel::Warning,
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
                            level: ChartDiagnosticLevel::Warning,
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
                        if rels_bytes.is_none() {
                            diagnostics.push(ChartDiagnostic {
                                level: ChartDiagnosticLevel::Warning,
                                message: format!(
                                    "missing chart userShapes relationships part: {rels_path}"
                                ),
                                part: Some(rels_path.clone()),
                                xpath: None,
                            });
                        }
                        Some(OpcPart {
                            path,
                            rels_path: Some(rels_path),
                            rels_bytes,
                            bytes: bytes.to_vec(),
                        })
                    }
                    None => {
                        diagnostics.push(ChartDiagnostic {
                            level: ChartDiagnosticLevel::Warning,
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
                                level: ChartDiagnosticLevel::Warning,
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
                                            level: ChartDiagnosticLevel::Warning,
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
                                level: ChartDiagnosticLevel::Warning,
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
                                level: ChartDiagnosticLevel::Warning,
                                message: format!(
                                    "failed to parse chartStyle part {}: {}",
                                    style_part.path,
                                    format_chart_style_error(&err)
                                ),
                                part: Some(style_part.path.clone()),
                                xpath: None,
                            }),
                        }
                        // Even if parsing fails, preserve the raw XML on the model for
                        // debugging/round-tripping.
                        if model.style_part.is_none() {
                            model.style_part = Some(ChartStylePartModel {
                                id: None,
                                raw_xml: String::from_utf8_lossy(&style_part.bytes).into_owned(),
                            });
                        }
                    }

                    if let Some(colors_part) = colors.as_ref() {
                        match parse_chart_color_style(&colors_part.bytes, &colors_part.path) {
                            Ok(colors_model) => model.colors_part = Some(colors_model),
                            Err(err) => diagnostics.push(ChartDiagnostic {
                                level: ChartDiagnosticLevel::Warning,
                                message: format!(
                                    "failed to parse chartColorStyle part {}: {}",
                                    colors_part.path,
                                    format_chart_color_style_error(&err)
                                ),
                                part: Some(colors_part.path.clone()),
                                xpath: None,
                            }),
                        }
                        if model.colors_part.is_none() {
                            model.colors_part = Some(ChartColorStylePartModel {
                                id: None,
                                colors: Vec::new(),
                                raw_xml: String::from_utf8_lossy(&colors_part.bytes).into_owned(),
                            });
                        }
                    }
                }

                chart_objects.push(ChartObject {
                    sheet_name: sheet_name.clone(),
                    sheet_part: sheet_part.clone(),
                    drawing_part: drawing_part.clone(),
                    drawing_rel_id: drawing_ref.rel_id,
                    drawing_object_id: drawing_ref.drawing_object_id,
                    drawing_object_name: drawing_ref.drawing_object_name,
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

            fn is_drawing_node(node: Node<'_, '_>) -> bool {
                node.is_element() && node.tag_name().name() == "drawing"
            }

            let drawing_rids: Vec<String> =
                crate::drawingml::anchor::descendants_selecting_alternate_content(
                    sheet_doc.root_element(),
                    is_drawing_node,
                    is_drawing_node,
                )
                .into_iter()
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
    if crate::ascii::contains_ignore_case(rel_type, "chartstyle") {
        return true;
    }
    // Fallback to filename heuristic for producers that omit the relationship type.
    crate::ascii::ends_with_ignore_case(rel_target, ".xml")
        && crate::ascii::contains_ignore_case(rel_target, "style")
        && !crate::ascii::ends_with_ignore_case(rel_target, "styles.xml")
}

fn is_chart_colors_relationship(rel_type: &str, rel_target: &str) -> bool {
    if crate::ascii::contains_ignore_case(rel_type, "chartcolorstyle") {
        return true;
    }
    crate::ascii::ends_with_ignore_case(rel_target, ".xml")
        && crate::ascii::contains_ignore_case(rel_target, "colors")
}

fn is_chart_user_shapes_relationship(rel_type: &str, rel_target: &str) -> bool {
    if crate::ascii::contains_ignore_case(rel_type, "chartusershapes") {
        return true;
    }
    // Fallback to filename heuristic for producers that omit the relationship type.
    if !crate::ascii::ends_with_ignore_case(rel_target, ".xml") {
        return false;
    }
    if crate::ascii::contains_ignore_case(rel_target, "usershapes") {
        return true;
    }
    // Only consider the final path component so a directory like `../drawings/` doesn't cause us
    // to misclassify unrelated `*.xml` parts as userShapes.
    let file_name = rel_target.rsplit('/').next().unwrap_or(rel_target);
    crate::ascii::contains_ignore_case(file_name, "drawing")
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
    fn chart_ex_kind_is_unknown(kind: &ChartKind) -> bool {
        let ChartKind::Unknown { name } = kind else {
            return false;
        };
        let tail = name.strip_prefix("ChartEx:").unwrap_or(name.as_str());
        let tail = tail.trim();
        tail.is_empty() || tail.eq_ignore_ascii_case("unknown")
    }

    fn normalize_chart_kind_hint(raw: &str) -> Option<String> {
        let raw = raw.trim();
        if raw.is_empty() {
            return None;
        }

        // Some producers include prefixes; keep only the local identifier.
        let raw = raw.split(':').last().unwrap_or(raw).trim();
        if raw.is_empty() {
            return None;
        }
        if raw == "missingPlotArea" || raw == "missingChartType" {
            return None;
        }

        let base = if raw
            .get(raw.len().saturating_sub("chart".len())..)
            .is_some_and(|tail| tail.eq_ignore_ascii_case("chart"))
        {
            &raw[..raw.len() - 5]
        } else {
            raw
        };
        if base.is_empty() {
            return None;
        }

        let mut chars = base.chars();
        let first = chars.next()?;
        let mut out = String::with_capacity(base.len());
        out.push(first.to_ascii_lowercase());
        out.push_str(chars.as_str());
        Some(out)
    }

    // Excel may store both a classic `c:chartSpace` part (`xl/charts/chartN.xml`) and a ChartEx
    // `cx:chartSpace` part (`xl/charts/chartExN.xml`) for the same chart. Our ChartEx parsing is
    // best-effort and can be minimal; when both models are available we merge them as follows:
    //
    // - Preserve ChartEx identity: keep the ChartEx `chart_kind` (typically `ChartEx:*`) and keep
    //   ChartEx diagnostics.
    // - Fill missing fields from the classic chartSpace model (series, axes, title/legend, plot
    //   area/layout/style) when ChartEx does not provide them.
    // - Never drop model diagnostics: append chartSpace diagnostics after ChartEx diagnostics.
    let chart_space_series_len = chart_space.series.len();
    let chart_ex_series_len = chart_ex.series.len();
    let chart_space_axes_len = chart_space.axes.len();
    let chart_ex_axes_len = chart_ex.axes.len();

    let mut chart_space = chart_space;
    let mut merged = chart_ex;

    // `cx:chartSpace` parts are sometimes minimal (title only), which means the ChartEx parser can
    // only return `ChartEx:unknown`. When the classic chartSpace model has a more specific chart
    // type (even if it's not yet a first-class `ChartKind`), propagate it into the merged model so
    // downstream consumers (like the legacy `extract_charts()` API) can still identify the chart
    // type.
    if chart_ex_kind_is_unknown(&merged.chart_kind) {
        if let ChartKind::Unknown { name } = &chart_space.chart_kind {
            if let Some(kind) = normalize_chart_kind_hint(name) {
                merged.chart_kind = ChartKind::Unknown {
                    name: format!("ChartEx:{kind}"),
                };
            }
        }
    }

    diagnostics.push(ChartDiagnostic {
        level: ChartDiagnosticLevel::Info,
        message: format!(
            "model.chart_kind: using ChartEx {chart_ex_kind:?} (chartSpace was {chart_space_kind:?})",
            chart_ex_kind = &merged.chart_kind,
            chart_space_kind = &chart_space.chart_kind,
        ),
        part: Some(chart_ex_part.to_string()),
        xpath: None,
    });

    // Preserve whichever diagnostics are available from both models. Keep ChartEx first.
    merged
        .diagnostics
        .extend(std::mem::take(&mut chart_space.diagnostics));

    // Series / axes: only fall back to chartSpace when ChartEx is missing them entirely.
    if merged.series.is_empty() && !chart_space.series.is_empty() {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: format!(
                "model.series: using chartSpace (ChartEx empty, chartSpace={chart_space_series_len})",
            ),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
        merged.series = std::mem::take(&mut chart_space.series);
    } else {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: format!(
                "model.series: using ChartEx (ChartEx={chart_ex_series_len}, chartSpace={chart_space_series_len})",
            ),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
    }

    if merged.axes.is_empty() && !chart_space.axes.is_empty() {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: format!(
                "model.axes: using chartSpace (ChartEx empty, chartSpace={chart_space_axes_len})",
            ),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
        merged.axes = std::mem::take(&mut chart_space.axes);
    } else {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: format!(
                "model.axes: using ChartEx (ChartEx={chart_ex_axes_len}, chartSpace={chart_space_axes_len})",
            ),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
    }

    // Title / legend: fall back to chartSpace when ChartEx is missing.
    if merged.title.is_none() && chart_space.title.is_some() {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: "model.title: using chartSpace".to_string(),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
        merged.title = chart_space.title.take();
    } else {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: "model.title: using ChartEx".to_string(),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
    }

    // Manual layout for text elements (title/legend) is often stored in the classic chartSpace
    // part even when the ChartEx part provides the actual title/legend content. Preserve that
    // layout information when ChartEx omits it so downstream rendering remains faithful.
    if let (Some(title), Some(chart_space_title)) =
        (merged.title.as_mut(), chart_space.title.as_ref())
    {
        if title.layout.is_none() {
            title.layout = chart_space_title.layout.clone();
        }
    }

    if merged.legend.is_none() && chart_space.legend.is_some() {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: "model.legend: using chartSpace".to_string(),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
        merged.legend = chart_space.legend.take();
    } else {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: "model.legend: using ChartEx".to_string(),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
    }

    if let (Some(legend), Some(chart_space_legend)) =
        (merged.legend.as_mut(), chart_space.legend.as_ref())
    {
        if legend.layout.is_none() {
            legend.layout = chart_space_legend.layout.clone();
        }
    }

    // Plot area: prefer chartSpace when ChartEx only provides an `Unknown` placeholder and the
    // chartSpace model has any non-"missing*" value.
    let chart_space_plot_area_present = match &chart_space.plot_area {
        PlotAreaModel::Unknown { name } => name != "missingPlotArea" && name != "missingChartType",
        _ => true,
    };

    if matches!(merged.plot_area, PlotAreaModel::Unknown { .. }) && chart_space_plot_area_present {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: "model.plot_area: using chartSpace".to_string(),
            part: Some(chart_space_part.to_string()),
            xpath: None,
        });
        merged.plot_area = std::mem::replace(
            &mut chart_space.plot_area,
            PlotAreaModel::Unknown {
                name: "missingPlotArea".to_string(),
            },
        );
    } else {
        diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Info,
            message: "model.plot_area: using ChartEx".to_string(),
            part: Some(chart_ex_part.to_string()),
            xpath: None,
        });
    }

    if merged.plot_area_layout.is_none() {
        merged.plot_area_layout = chart_space.plot_area_layout.take();
    }

    if merged.style_id.is_none() {
        merged.style_id = chart_space.style_id.take();
    }

    if merged.rounded_corners.is_none() {
        merged.rounded_corners = chart_space.rounded_corners.take();
    }

    if merged.disp_blanks_as.is_none() {
        merged.disp_blanks_as = chart_space.disp_blanks_as.take();
    }

    if merged.plot_vis_only.is_none() {
        merged.plot_vis_only = chart_space.plot_vis_only.take();
    }

    if merged.chart_area_style.is_none() {
        merged.chart_area_style = chart_space.chart_area_style.take();
    }

    if merged.plot_area_style.is_none() {
        merged.plot_area_style = chart_space.plot_area_style.take();
    }

    if merged.chart_space_ext_lst_xml.is_none() {
        merged.chart_space_ext_lst_xml = chart_space.chart_space_ext_lst_xml.take();
    }

    if merged.chart_ext_lst_xml.is_none() {
        merged.chart_ext_lst_xml = chart_space.chart_ext_lst_xml.take();
    }

    if merged.plot_area_ext_lst_xml.is_none() {
        merged.plot_area_ext_lst_xml = chart_space.plot_area_ext_lst_xml.take();
    }

    // External workbook links: prefer the classic chartSpace value when present (that's where
    // Excel usually stores it), otherwise fall back to ChartEx.
    let chart_ex_rel_id = merged.external_data_rel_id.clone();
    let chart_ex_auto_update = merged.external_data_auto_update;

    match (
        chart_space.external_data_rel_id.as_deref(),
        merged.external_data_rel_id.as_deref(),
    ) {
        (Some(chart_space_rid), None) => {
            diagnostics.push(ChartDiagnostic {
                level: ChartDiagnosticLevel::Info,
                message: format!(
                    "model.external_data_rel_id: using chartSpace {chart_space_rid:?} (ChartEx was None)",
                ),
                part: Some(chart_space_part.to_string()),
                xpath: None,
            });
            merged.external_data_rel_id = chart_space.external_data_rel_id.take();
        }
        (Some(chart_space_rid), Some(chart_ex_rid)) if chart_space_rid != chart_ex_rid => {
            diagnostics.push(ChartDiagnostic {
                level: ChartDiagnosticLevel::Warning,
                message: format!(
                    "model.external_data_rel_id: chartSpace {chart_space_rid:?} differs from ChartEx {chart_ex_rid:?}; using chartSpace",
                ),
                part: Some(chart_space_part.to_string()),
                xpath: None,
            });
            merged.external_data_rel_id = chart_space.external_data_rel_id.take();
        }
        _ => {}
    }

    if let Some(chart_space_auto) = chart_space.external_data_auto_update {
        if let Some(chart_ex_auto) = merged.external_data_auto_update {
            if chart_space_auto != chart_ex_auto {
                diagnostics.push(ChartDiagnostic {
                    level: ChartDiagnosticLevel::Warning,
                    message: format!(
                        "model.external_data_auto_update: chartSpace {chart_space_auto:?} differs from ChartEx {chart_ex_auto:?}; using chartSpace",
                    ),
                    part: Some(chart_space_part.to_string()),
                    xpath: None,
                });
            }
        }
        merged.external_data_auto_update = Some(chart_space_auto);
    } else if chart_ex_auto_update.is_some() {
        // Only keep the ChartEx autoUpdate value if the relationship id still matches after
        // resolving any chartSpace vs ChartEx externalData relationship conflicts.
        let rel_ids_match = match (&merged.external_data_rel_id, &chart_ex_rel_id) {
            (Some(a), Some(b)) => a == b,
            (Some(_), None) => true,
            _ => false,
        };
        if !rel_ids_match {
            diagnostics.push(ChartDiagnostic {
                level: ChartDiagnosticLevel::Warning,
                message: "model.external_data_auto_update: ignoring ChartEx autoUpdate due to externalData rel id mismatch".to_string(),
                part: Some(chart_ex_part.to_string()),
                xpath: None,
            });
            merged.external_data_auto_update = None;
        }
    }

    merged
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::charts::{
        LegendModel, LegendPosition, ManualLayoutModel, PlotAreaModel, TextModel,
    };

    fn minimal_model(rel_id: Option<&str>, auto_update: Option<bool>) -> ChartModel {
        ChartModel {
            chart_kind: ChartKind::Unknown {
                name: "unknown".to_string(),
            },
            title: None,
            legend: None,
            plot_area: PlotAreaModel::Unknown {
                name: "unknown".to_string(),
            },
            plot_area_layout: None,
            axes: Vec::new(),
            series: Vec::new(),
            style_id: None,
            rounded_corners: None,
            disp_blanks_as: None,
            plot_vis_only: None,
            style_part: None,
            colors_part: None,
            chart_area_style: None,
            plot_area_style: None,
            external_data_rel_id: rel_id.map(str::to_string),
            external_data_auto_update: auto_update,
            chart_space_ext_lst_xml: None,
            chart_ext_lst_xml: None,
            plot_area_ext_lst_xml: None,
            diagnostics: Vec::new(),
        }
    }

    #[test]
    fn merge_chart_models_preserves_title_layout_from_chart_space_when_chart_ex_omits_it() {
        let mut chart_space = minimal_model(None, None);
        let mut chart_ex = minimal_model(None, None);

        chart_space.chart_kind = ChartKind::Bar;
        let mut chart_space_title = TextModel::plain("chartSpace title");
        chart_space_title.layout = Some(ManualLayoutModel {
            x: Some(0.25),
            ..Default::default()
        });
        chart_space.title = Some(chart_space_title);

        chart_ex.chart_kind = ChartKind::Unknown {
            name: "ChartEx:histogram".to_string(),
        };
        chart_ex.title = Some(TextModel::plain("ChartEx title"));

        let mut diagnostics = Vec::new();
        let merged = merge_chart_models(
            chart_space,
            chart_ex,
            "xl/charts/chart1.xml",
            "xl/charts/chartEx1.xml",
            &mut diagnostics,
        );

        let merged_title = merged.title.expect("title present");
        assert_eq!(merged_title.rich_text.plain_text(), "ChartEx title");
        assert_eq!(merged_title.layout.as_ref().and_then(|l| l.x), Some(0.25));
    }

    #[test]
    fn merge_chart_models_preserves_legend_layout_from_chart_space_when_chart_ex_omits_it() {
        let mut chart_space = minimal_model(None, None);
        let mut chart_ex = minimal_model(None, None);

        chart_space.legend = Some(LegendModel {
            position: LegendPosition::Right,
            overlay: false,
            text_style: None,
            style: None,
            layout: Some(ManualLayoutModel {
                y: Some(0.5),
                ..Default::default()
            }),
        });

        chart_ex.legend = Some(LegendModel {
            position: LegendPosition::Left,
            overlay: true,
            text_style: None,
            style: None,
            layout: None,
        });

        let mut diagnostics = Vec::new();
        let merged = merge_chart_models(
            chart_space,
            chart_ex,
            "xl/charts/chart1.xml",
            "xl/charts/chartEx1.xml",
            &mut diagnostics,
        );

        let merged_legend = merged.legend.expect("legend present");
        assert_eq!(merged_legend.position, LegendPosition::Left);
        assert_eq!(merged_legend.layout.as_ref().and_then(|l| l.y), Some(0.5));
    }

    #[test]
    fn merge_chart_models_external_data_falls_back_to_chart_ex() {
        let chart_space = minimal_model(None, None);
        let chart_ex = minimal_model(Some("rId9"), Some(true));
        let mut diagnostics = Vec::new();

        let merged = merge_chart_models(
            chart_space,
            chart_ex,
            "xl/charts/chart1.xml",
            "xl/charts/chartEx1.xml",
            &mut diagnostics,
        );

        assert_eq!(merged.external_data_rel_id.as_deref(), Some("rId9"));
        assert_eq!(merged.external_data_auto_update, Some(true));
    }

    #[test]
    fn merge_chart_models_external_data_prefers_chart_space_on_conflict() {
        let chart_space = minimal_model(Some("rId1"), Some(false));
        let chart_ex = minimal_model(Some("rId2"), Some(true));
        let mut diagnostics = Vec::new();

        let merged = merge_chart_models(
            chart_space,
            chart_ex,
            "xl/charts/chart1.xml",
            "xl/charts/chartEx1.xml",
            &mut diagnostics,
        );

        assert_eq!(merged.external_data_rel_id.as_deref(), Some("rId1"));
        assert_eq!(merged.external_data_auto_update, Some(false));
    }

    #[test]
    fn merge_chart_models_infers_chart_ex_kind_from_chart_space_when_chart_ex_unknown() {
        let mut chart_space = minimal_model(None, None);
        chart_space.chart_kind = ChartKind::Unknown {
            name: "waterfallChart".to_string(),
        };

        let mut chart_ex = minimal_model(None, None);
        chart_ex.chart_kind = ChartKind::Unknown {
            name: "ChartEx:unknown".to_string(),
        };

        let mut diagnostics = Vec::new();
        let merged = merge_chart_models(
            chart_space,
            chart_ex,
            "xl/charts/chart1.xml",
            "xl/charts/chartEx1.xml",
            &mut diagnostics,
        );

        assert_eq!(
            merged.chart_kind,
            ChartKind::Unknown {
                name: "ChartEx:waterfall".to_string()
            }
        );
    }

    #[test]
    fn merge_chart_models_does_not_override_chart_ex_kind_when_chart_space_is_known() {
        let mut chart_space = minimal_model(None, None);
        chart_space.chart_kind = ChartKind::Bar;

        let mut chart_ex = minimal_model(None, None);
        chart_ex.chart_kind = ChartKind::Unknown {
            name: "ChartEx:unknown".to_string(),
        };

        let mut diagnostics = Vec::new();
        let merged = merge_chart_models(
            chart_space,
            chart_ex,
            "xl/charts/chart1.xml",
            "xl/charts/chartEx1.xml",
            &mut diagnostics,
        );

        assert_eq!(
            merged.chart_kind,
            ChartKind::Unknown {
                name: "ChartEx:unknown".to_string()
            }
        );
    }

    #[test]
    fn merge_chart_models_preserves_and_concatenates_model_diagnostics() {
        let mut chart_space = minimal_model(None, None);
        let mut chart_ex = minimal_model(None, None);

        chart_ex.chart_kind = ChartKind::Unknown {
            name: "ChartEx:histogram".to_string(),
        };
        chart_ex.diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Warning,
            message: "chartEx diagnostic".to_string(),
            part: None,
            xpath: None,
        });
        chart_space.diagnostics.push(ChartDiagnostic {
            level: ChartDiagnosticLevel::Warning,
            message: "chartSpace diagnostic".to_string(),
            part: None,
            xpath: None,
        });

        let mut diagnostics = Vec::new();
        let merged = merge_chart_models(
            chart_space,
            chart_ex,
            "xl/charts/chart1.xml",
            "xl/charts/chartEx1.xml",
            &mut diagnostics,
        );

        assert_eq!(merged.diagnostics.len(), 2);
        assert_eq!(merged.diagnostics[0].message, "chartEx diagnostic");
        assert_eq!(merged.diagnostics[1].message, "chartSpace diagnostic");
    }
}

fn normalize_relationship_target(target: &str) -> String {
    // Relationship targets are URIs. For in-package parts, the `#fragment` and `?query` portions
    // are not part of the OPC part name and must be ignored when mapping to ZIP entry names.
    let target = target.split_once('#').map(|(t, _)| t).unwrap_or(target);
    let target = target.split_once('?').map(|(t, _)| t).unwrap_or(target);
    let mut out = if target.contains('\\') {
        target.replace('\\', "/")
    } else {
        target.to_string()
    };
    while let Some(stripped) = out.strip_prefix("./") {
        out = stripped.to_string();
    }
    out
}

fn is_external_target_mode(target_mode: Option<&str>) -> bool {
    target_mode.is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
}

fn chart_anchor_from_drawing_anchor(anchor: DrawingAnchor) -> ChartAnchor {
    match anchor {
        DrawingAnchor::TwoCell { from, to } => ChartAnchor::TwoCell {
            from_col: from.cell.col,
            from_row: from.cell.row,
            from_col_off_emu: from.offset.x_emu,
            from_row_off_emu: from.offset.y_emu,
            to_col: to.cell.col,
            to_row: to.cell.row,
            to_col_off_emu: to.offset.x_emu,
            to_row_off_emu: to.offset.y_emu,
        },
        DrawingAnchor::OneCell { from, ext } => ChartAnchor::OneCell {
            from_col: from.cell.col,
            from_row: from.cell.row,
            from_col_off_emu: from.offset.x_emu,
            from_row_off_emu: from.offset.y_emu,
            cx_emu: ext.cx,
            cy_emu: ext.cy,
        },
        DrawingAnchor::Absolute { pos, ext } => ChartAnchor::Absolute {
            x_emu: pos.x_emu,
            y_emu: pos.y_emu,
            cx_emu: ext.cx,
            cy_emu: ext.cy,
        },
    }
}
