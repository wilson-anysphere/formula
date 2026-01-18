use formula_model::{
    extract_a1_references, parse_argb_hex_color, parse_sqref, CellIsOperator, CfRule, CfRuleKind,
    CfRuleSchema, Cfvo, CfvoType, ColorScaleRule, DataBarDirection, DataBarRule, IconSet, IconSetRule,
    TopBottomKind, TopBottomRule, UniqueDuplicateRule,
};
use roxmltree::Document;
use std::collections::HashMap;

#[derive(Debug, thiserror::Error)]
pub enum ConditionalFormattingError {
    #[error("xml parse error: {0}")]
    Xml(#[from] roxmltree::Error),
    #[error("missing sqref for conditionalFormatting")]
    MissingSqref,
}

#[derive(Clone, Debug)]
pub struct RawConditionalFormattingBlock {
    pub schema: CfRuleSchema,
    pub xml: String,
}

#[derive(Clone, Debug, Default)]
pub struct ParsedConditionalFormatting {
    pub rules: Vec<CfRule>,
    pub raw_blocks: Vec<RawConditionalFormattingBlock>,
}

pub fn parse_worksheet_conditional_formatting(xml: &str) -> Result<ParsedConditionalFormatting, ConditionalFormattingError> {
    let doc = Document::parse(xml)?;
    let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let x14_ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    let xm_ns = "http://schemas.microsoft.com/office/excel/2006/main";

    let mut parsed = ParsedConditionalFormatting::default();

    let mut base_rules_by_id: HashMap<String, usize> = HashMap::new();

    for cf in doc.descendants().filter(|n| {
        n.is_element() && n.tag_name().name() == "conditionalFormatting" && n.tag_name().namespace() == Some(main_ns)
    }) {
        let range = cf.range();
        parsed.raw_blocks.push(RawConditionalFormattingBlock {
            schema: CfRuleSchema::Office2007,
            xml: xml[range].to_string(),
        });

        let sqref = cf.attribute("sqref").ok_or(ConditionalFormattingError::MissingSqref)?;
        let applies_to = parse_sqref(sqref).unwrap_or_default();

        for rule_node in cf.children().filter(|n| n.is_element() && n.tag_name().name() == "cfRule") {
            let raw_xml = xml[rule_node.range()].to_string();
            let priority = rule_node
                .attribute("priority")
                .and_then(|p| p.parse::<u32>().ok())
                .unwrap_or(u32::MAX);
            let dxf_id = rule_node.attribute("dxfId").and_then(|id| id.parse::<u32>().ok());
            let stop_if_true = rule_node
                .attribute("stopIfTrue")
                .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                .unwrap_or(false);
            let id = rule_node.attribute("id").map(|s| s.to_string());

            let kind = parse_cf_rule_kind(rule_node, CfRuleSchema::Office2007, main_ns, x14_ns, &raw_xml);
            let dependencies = compute_dependencies(&applies_to, &kind);

            let mut rule = CfRule {
                schema: CfRuleSchema::Office2007,
                id: id.clone(),
                priority,
                applies_to: applies_to.clone(),
                dxf_id,
                stop_if_true,
                kind,
                dependencies,
            };

            if let Some(id) = id {
                base_rules_by_id.insert(id, parsed.rules.len());
            }

            // If this is a dataBar/iconSet/etc and has x14 extension, we might upgrade it later.
            rule.schema = CfRuleSchema::Office2007;
            parsed.rules.push(rule);
        }
    }

    // Parse x14 conditionalFormattings, and merge onto base rules when possible.
    for cf14 in doc.descendants().filter(|n| {
        n.is_element() && n.tag_name().name() == "conditionalFormattings" && n.tag_name().namespace() == Some(x14_ns)
    }) {
        let range = cf14.range();
        parsed.raw_blocks.push(RawConditionalFormattingBlock {
            schema: CfRuleSchema::X14,
            xml: xml[range].to_string(),
        });

        for block in cf14.children().filter(|n| n.is_element() && n.tag_name().name() == "conditionalFormatting") {
            let sqref_text = block
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "sqref" && n.tag_name().namespace() == Some(xm_ns))
                .and_then(|n| n.text())
                .unwrap_or("");
            let applies_to = parse_sqref(sqref_text).unwrap_or_default();

            for rule_node in block
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "cfRule")
            {
                let raw_xml = xml[rule_node.range()].to_string();
                let id = rule_node.attribute("id").map(|s| s.to_string());
                let priority = rule_node
                    .attribute("priority")
                    .and_then(|p| p.parse::<u32>().ok())
                    .unwrap_or(u32::MAX);

                let kind = parse_cf_rule_kind(rule_node, CfRuleSchema::X14, main_ns, x14_ns, &raw_xml);
                let dependencies = compute_dependencies(&applies_to, &kind);

                let x14_rule = CfRule {
                    schema: CfRuleSchema::X14,
                    id: id.clone(),
                    priority,
                    applies_to: applies_to.clone(),
                    dxf_id: None,
                    stop_if_true: false,
                    kind,
                    dependencies,
                };

                if let Some(id) = id.as_deref() {
                    if let Some(&idx) = base_rules_by_id.get(id) {
                        // Merge x14 extensions into base rule.
                        let base = &mut parsed.rules[idx];
                        merge_x14_into_base(base, &x14_rule);
                        base.schema = CfRuleSchema::X14;
                        continue;
                    }
                }

                parsed.rules.push(x14_rule);
            }
        }
    }

    Ok(parsed)
}

fn parse_cf_rule_kind(
    rule_node: roxmltree::Node<'_, '_>,
    schema: CfRuleSchema,
    main_ns: &str,
    x14_ns: &str,
    raw_xml: &str,
) -> CfRuleKind {
    let type_name = rule_node.attribute("type").unwrap_or("");
    match type_name {
        "cellIs" => {
            let operator = match rule_node.attribute("operator").unwrap_or("equal") {
                "greaterThan" => CellIsOperator::GreaterThan,
                "greaterThanOrEqual" => CellIsOperator::GreaterThanOrEqual,
                "lessThan" => CellIsOperator::LessThan,
                "lessThanOrEqual" => CellIsOperator::LessThanOrEqual,
                "notEqual" => CellIsOperator::NotEqual,
                "between" => CellIsOperator::Between,
                "notBetween" => CellIsOperator::NotBetween,
                _ => CellIsOperator::Equal,
            };
            let formulas: Vec<String> = rule_node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "formula")
                .filter_map(|n| n.text().map(|t| t.to_string()))
                .collect();
            CfRuleKind::CellIs { operator, formulas }
        }
        "expression" => {
            let formula = rule_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "formula")
                .and_then(|n| n.text())
                .unwrap_or("")
                .to_string();
            CfRuleKind::Expression { formula }
        }
        "dataBar" => {
            if schema == CfRuleSchema::X14 {
                parse_x14_data_bar(rule_node, x14_ns).unwrap_or_else(|| CfRuleKind::Unsupported {
                    type_name: Some(type_name.to_string()),
                    raw_xml: raw_xml.to_string(),
                })
            } else {
                parse_data_bar(rule_node, main_ns).unwrap_or_else(|| CfRuleKind::Unsupported {
                    type_name: Some(type_name.to_string()),
                    raw_xml: raw_xml.to_string(),
                })
            }
        }
        "colorScale" => parse_color_scale(rule_node, main_ns).unwrap_or_else(|| CfRuleKind::Unsupported {
            type_name: Some(type_name.to_string()),
            raw_xml: raw_xml.to_string(),
        }),
        "iconSet" => parse_icon_set(rule_node, main_ns).unwrap_or_else(|| CfRuleKind::Unsupported {
            type_name: Some(type_name.to_string()),
            raw_xml: raw_xml.to_string(),
        }),
        "top10" => parse_top10(rule_node).unwrap_or_else(|| CfRuleKind::Unsupported {
            type_name: Some(type_name.to_string()),
            raw_xml: raw_xml.to_string(),
        }),
        "uniqueValues" => CfRuleKind::UniqueDuplicate(UniqueDuplicateRule { unique: true }),
        "duplicateValues" => CfRuleKind::UniqueDuplicate(UniqueDuplicateRule { unique: false }),
        _ => CfRuleKind::Unsupported {
            type_name: if type_name.is_empty() { None } else { Some(type_name.to_string()) },
            raw_xml: raw_xml.to_string(),
        },
    }
}

fn parse_data_bar(rule_node: roxmltree::Node<'_, '_>, main_ns: &str) -> Option<CfRuleKind> {
    let data_bar = rule_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dataBar" && n.tag_name().namespace() == Some(main_ns))?;
    let mut cfvos = data_bar
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "cfvo")
        .map(parse_cfvo)
        .collect::<Vec<_>>();
    if cfvos.len() < 2 {
        return None;
    }
    let max = cfvos.pop()?;
    let min = cfvos.pop()?;
    let color = data_bar
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "color")
        .and_then(|c| c.attribute("rgb"))
        .and_then(parse_argb_hex_color);

    Some(CfRuleKind::DataBar(DataBarRule {
        min,
        max,
        color,
        min_length: None,
        max_length: None,
        gradient: None,
        negative_fill_color: None,
        axis_color: None,
        direction: None,
    }))
}

fn parse_x14_data_bar(rule_node: roxmltree::Node<'_, '_>, x14_ns: &str) -> Option<CfRuleKind> {
    let data_bar = rule_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dataBar" && n.tag_name().namespace() == Some(x14_ns))?;
    let mut cfvos = data_bar
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "cfvo")
        .map(parse_cfvo)
        .collect::<Vec<_>>();
    if cfvos.len() < 2 {
        return None;
    }
    let max = cfvos.pop()?;
    let min = cfvos.pop()?;

    let min_length = data_bar.attribute("minLength").and_then(|v| v.parse::<u8>().ok());
    let max_length = data_bar.attribute("maxLength").and_then(|v| v.parse::<u8>().ok());
    let gradient = data_bar
        .attribute("gradient")
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
    let direction = data_bar.attribute("direction").and_then(|v| {
        if v.eq_ignore_ascii_case("leftToRight") {
            Some(DataBarDirection::LeftToRight)
        } else if v.eq_ignore_ascii_case("rightToLeft") {
            Some(DataBarDirection::RightToLeft)
        } else if v.eq_ignore_ascii_case("context") {
            Some(DataBarDirection::Context)
        } else {
            None
        }
    });

    let negative_fill_color = data_bar
        .children()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "negativeFillColor"
                && n.tag_name().namespace() == Some(x14_ns)
        })
        .and_then(|c| c.attribute("rgb"))
        .and_then(parse_argb_hex_color);
    let axis_color = data_bar
        .children()
        .find(|n| {
            n.is_element() && n.tag_name().name() == "axisColor" && n.tag_name().namespace() == Some(x14_ns)
        })
        .and_then(|c| c.attribute("rgb"))
        .and_then(parse_argb_hex_color);

    // x14 extended schema typically omits the positive fill color; it remains in the base cfRule.
    Some(CfRuleKind::DataBar(DataBarRule {
        min,
        max,
        color: None,
        min_length,
        max_length,
        gradient,
        negative_fill_color,
        axis_color,
        direction,
    }))
}

fn parse_color_scale(rule_node: roxmltree::Node<'_, '_>, main_ns: &str) -> Option<CfRuleKind> {
    let cs = rule_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "colorScale" && n.tag_name().namespace() == Some(main_ns))?;
    let cfvos = cs
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "cfvo")
        .map(parse_cfvo)
        .collect::<Vec<_>>();
    let colors = cs
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "color")
        .filter_map(|c| c.attribute("rgb").and_then(parse_argb_hex_color))
        .collect::<Vec<_>>();
    if colors.len() < 2 {
        return None;
    }
    Some(CfRuleKind::ColorScale(ColorScaleRule { cfvos, colors }))
}

fn parse_icon_set(rule_node: roxmltree::Node<'_, '_>, main_ns: &str) -> Option<CfRuleKind> {
    let is = rule_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "iconSet" && n.tag_name().namespace() == Some(main_ns))?;
    let set_name = is.attribute("iconSet").unwrap_or("3Arrows");
    let set = match set_name {
        "3Arrows" => IconSet::ThreeArrows,
        "3TrafficLights1" => IconSet::ThreeTrafficLights1,
        "3TrafficLights2" => IconSet::ThreeTrafficLights2,
        "3Flags" => IconSet::ThreeFlags,
        "3Symbols" => IconSet::ThreeSymbols,
        "3Symbols2" => IconSet::ThreeSymbols2,
        "4Arrows" => IconSet::FourArrows,
        "4ArrowsGray" => IconSet::FourArrowsGray,
        "5Arrows" => IconSet::FiveArrows,
        "5ArrowsGray" => IconSet::FiveArrowsGray,
        "5Quarters" => IconSet::FiveQuarters,
        _ => IconSet::ThreeArrows,
    };
    let cfvos = is
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "cfvo")
        .map(parse_cfvo)
        .collect::<Vec<_>>();
    let show_value = is
        .attribute("showValue")
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(true);
    let reverse = is
        .attribute("reverse")
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false);
    Some(CfRuleKind::IconSet(IconSetRule {
        set,
        cfvos,
        show_value,
        reverse,
    }))
}

fn parse_top10(rule_node: roxmltree::Node<'_, '_>) -> Option<CfRuleKind> {
    let rank = rule_node.attribute("rank").and_then(|v| v.parse::<u32>().ok()).unwrap_or(10);
    let percent = rule_node
        .attribute("percent")
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false);
    let bottom = rule_node
        .attribute("bottom")
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false);
    let kind = if bottom { TopBottomKind::Bottom } else { TopBottomKind::Top };
    Some(CfRuleKind::TopBottom(TopBottomRule { kind, rank, percent }))
}

fn parse_cfvo(node: roxmltree::Node<'_, '_>) -> Cfvo {
    let type_ = match node.attribute("type").unwrap_or("min") {
        "min" => CfvoType::Min,
        "max" => CfvoType::Max,
        "num" => CfvoType::Number,
        "percent" => CfvoType::Percent,
        "percentile" => CfvoType::Percentile,
        "formula" => CfvoType::Formula,
        "autoMin" => CfvoType::AutoMin,
        "autoMax" => CfvoType::AutoMax,
        other => {
            let _ = other;
            CfvoType::Number
        }
    };
    let value = node.attribute("val").map(|s| s.to_string());
    Cfvo { type_, value }
}

fn merge_x14_into_base(base: &mut CfRule, ext: &CfRule) {
    match (&mut base.kind, &ext.kind) {
        (CfRuleKind::DataBar(base_db), CfRuleKind::DataBar(ext_db)) => {
            // Prefer values already present in the base cfRule, and fill in missing ones from x14.
            base_db.color = base_db.color.or(ext_db.color);
            base_db.min_length = base_db.min_length.or(ext_db.min_length);
            base_db.max_length = base_db.max_length.or(ext_db.max_length);
            base_db.gradient = base_db.gradient.or(ext_db.gradient);
            base_db.negative_fill_color = base_db.negative_fill_color.or(ext_db.negative_fill_color);
            base_db.axis_color = base_db.axis_color.or(ext_db.axis_color);
            base_db.direction = base_db.direction.or(ext_db.direction);
        }
        _ => {}
    }
}

fn compute_dependencies(applies_to: &[formula_model::Range], kind: &CfRuleKind) -> Vec<formula_model::Range> {
    let mut deps = Vec::new();
    deps.extend_from_slice(applies_to);
    match kind {
        CfRuleKind::CellIs { formulas, .. } => {
            for f in formulas {
                deps.extend(extract_a1_references(f));
            }
        }
        CfRuleKind::Expression { formula } => {
            deps.extend(extract_a1_references(formula));
        }
        CfRuleKind::DataBar(db) => {
            if db.min.type_ == CfvoType::Formula {
                deps.extend(extract_a1_references(db.min.value.as_deref().unwrap_or("")));
            }
            if db.max.type_ == CfvoType::Formula {
                deps.extend(extract_a1_references(db.max.value.as_deref().unwrap_or("")));
            }
        }
        CfRuleKind::ColorScale(cs) => {
            for cfvo in &cs.cfvos {
                if cfvo.type_ == CfvoType::Formula {
                    deps.extend(extract_a1_references(cfvo.value.as_deref().unwrap_or("")));
                }
            }
        }
        CfRuleKind::IconSet(is) => {
            for cfvo in &is.cfvos {
                if cfvo.type_ == CfvoType::Formula {
                    deps.extend(extract_a1_references(cfvo.value.as_deref().unwrap_or("")));
                }
            }
        }
        _ => {}
    }
    deps.sort_by_key(|r| (r.start.row, r.start.col, r.end.row, r.end.col));
    deps.dedup();
    deps
}
