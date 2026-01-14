use std::collections::BTreeMap;

use formula_model::{
    CellIsOperator, CfRule, CfRuleKind, CfRuleSchema, Cfvo, CfvoType, Color, ColorScaleRule,
    DataBarRule, IconSet, IconSetRule, Range, TopBottomKind, TopBottomRule, UniqueDuplicateRule,
};

const X14_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS: &str = "http://schemas.microsoft.com/office/excel/2006/main";
const X14_CONDITIONAL_FORMATTING_URI: &str = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";

/// Serialize conditional formatting rules to SpreadsheetML worksheet XML fragments.
///
/// This returns a string containing one or more `<conditionalFormatting>` blocks (SpreadsheetML
/// 2006 schema), followed by an `<extLst>` block when any rules require x14 extensions.
///
/// Formula strings in `<cfRule><formula>` and `<cfvo type="formula">` are normalized into the
/// OOXML file form (no leading `'='`, `_xlfn.` prefixes applied for forward-compatible functions).
#[must_use]
pub fn write_conditional_formatting_xml(rules: &[CfRule]) -> String {
    if rules.is_empty() {
        return String::new();
    }

    // Group base conditional formatting rules by `sqref` (SpreadsheetML requires
    // `<conditionalFormatting sqref="...">` containers).
    let mut groups: BTreeMap<String, Vec<&CfRule>> = BTreeMap::new();
    for rule in rules {
        groups
            .entry(sqref_for_ranges(&rule.applies_to))
            .or_default()
            .push(rule);
    }

    let mut out = String::new();

    for (sqref, mut group) in groups {
        group.sort_by_key(|r| r.priority);
        out.push_str(r#"<conditionalFormatting sqref=""#);
        out.push_str(&escape_xml(&sqref));
        out.push_str(r#"">"#);
        for rule in group {
            out.push_str(&write_base_cf_rule(rule));
        }
        out.push_str("</conditionalFormatting>");
    }

    if let Some(ext_lst) = write_x14_ext_lst(rules) {
        out.push_str(&ext_lst);
    }

    out
}

fn write_base_cf_rule(rule: &CfRule) -> String {
    if let CfRuleKind::Unsupported { raw_xml, .. } = &rule.kind {
        // Best-effort: pass through the original raw rule XML.
        return raw_xml.clone();
    }

    let mut attrs = String::new();
    attrs.push_str(r#" priority=""#);
    attrs.push_str(&escape_xml(&rule.priority.to_string()));
    attrs.push('"');

    if let Some(dxf_id) = rule.dxf_id {
        attrs.push_str(r#" dxfId=""#);
        attrs.push_str(&escape_xml(&dxf_id.to_string()));
        attrs.push('"');
    }
    if rule.stop_if_true {
        attrs.push_str(r#" stopIfTrue="1""#);
    }
    if let Some(id) = &rule.id {
        attrs.push_str(r#" id=""#);
        attrs.push_str(&escape_xml(id));
        attrs.push('"');
    }

    match &rule.kind {
        CfRuleKind::CellIs { operator, formulas } => {
            let mut s = String::new();
            s.push_str(r#"<cfRule type="cellIs""#);
            s.push_str(&attrs);
            s.push_str(r#" operator=""#);
            s.push_str(cell_is_operator_attr(*operator));
            s.push_str(r#"">"#);
            for f in formulas {
                s.push_str("<formula>");
                s.push_str(&escape_xml(&normalize_cf_formula(f)));
                s.push_str("</formula>");
            }
            s.push_str("</cfRule>");
            s
        }
        CfRuleKind::Expression { formula } => {
            let mut s = String::new();
            s.push_str(r#"<cfRule type="expression""#);
            s.push_str(&attrs);
            s.push('>');
            s.push_str("<formula>");
            s.push_str(&escape_xml(&normalize_cf_formula(formula)));
            s.push_str("</formula>");
            s.push_str("</cfRule>");
            s
        }
        CfRuleKind::DataBar(rule) => {
            let mut s = String::new();
            s.push_str(r#"<cfRule type="dataBar""#);
            s.push_str(&attrs);
            s.push('>');
            s.push_str(&write_base_data_bar(rule));
            s.push_str("</cfRule>");
            s
        }
        CfRuleKind::ColorScale(rule) => {
            let mut s = String::new();
            s.push_str(r#"<cfRule type="colorScale""#);
            s.push_str(&attrs);
            s.push('>');
            s.push_str(&write_base_color_scale(rule));
            s.push_str("</cfRule>");
            s
        }
        CfRuleKind::IconSet(rule) => {
            let mut s = String::new();
            s.push_str(r#"<cfRule type="iconSet""#);
            s.push_str(&attrs);
            s.push('>');
            s.push_str(&write_base_icon_set(rule));
            s.push_str("</cfRule>");
            s
        }
        CfRuleKind::TopBottom(rule) => write_base_top10(rule, &attrs),
        CfRuleKind::UniqueDuplicate(rule) => write_base_unique_duplicate(rule, &attrs),
        CfRuleKind::Unsupported { .. } => unreachable!("handled above"),
    }
}

fn write_base_data_bar(rule: &DataBarRule) -> String {
    let mut s = String::new();
    s.push_str("<dataBar>");
    s.push_str(&write_cfvo(&rule.min, "cfvo"));
    s.push_str(&write_cfvo(&rule.max, "cfvo"));
    if let Some(rgb) = color_rgb(rule.color) {
        s.push_str(r#"<color rgb=""#);
        s.push_str(&rgb);
        s.push_str(r#""/>"#);
    }
    s.push_str("</dataBar>");
    s
}

fn write_base_color_scale(rule: &ColorScaleRule) -> String {
    let mut s = String::new();
    s.push_str("<colorScale>");
    for cfvo in &rule.cfvos {
        s.push_str(&write_cfvo(cfvo, "cfvo"));
    }
    for color in &rule.colors {
        if let Some(rgb) = color_rgb(Some(*color)) {
            s.push_str(r#"<color rgb=""#);
            s.push_str(&rgb);
            s.push_str(r#""/>"#);
        }
    }
    s.push_str("</colorScale>");
    s
}

fn write_base_icon_set(rule: &IconSetRule) -> String {
    let mut s = String::new();
    s.push_str(r#"<iconSet iconSet=""#);
    s.push_str(icon_set_attr(rule.set));
    s.push('"');
    if !rule.show_value {
        s.push_str(r#" showValue="0""#);
    }
    if rule.reverse {
        s.push_str(r#" reverse="1""#);
    }
    s.push('>');
    for cfvo in &rule.cfvos {
        s.push_str(&write_cfvo(cfvo, "cfvo"));
    }
    s.push_str("</iconSet>");
    s
}

fn write_base_top10(rule: &TopBottomRule, base_attrs: &str) -> String {
    let mut s = String::new();
    s.push_str(r#"<cfRule type="top10""#);
    s.push_str(base_attrs);

    s.push_str(r#" rank=""#);
    s.push_str(&rule.rank.to_string());
    s.push('"');

    if rule.percent {
        s.push_str(r#" percent="1""#);
    }
    if matches!(rule.kind, TopBottomKind::Bottom) {
        s.push_str(r#" bottom="1""#);
    }
    s.push_str("/>");
    s
}

fn write_base_unique_duplicate(rule: &UniqueDuplicateRule, base_attrs: &str) -> String {
    let type_name = if rule.unique {
        "uniqueValues"
    } else {
        "duplicateValues"
    };
    let mut s = String::new();
    s.push_str(r#"<cfRule type=""#);
    s.push_str(type_name);
    s.push('"');
    s.push_str(base_attrs);
    s.push_str("/>");
    s
}

fn write_x14_ext_lst(rules: &[CfRule]) -> Option<String> {
    let x14_rules: Vec<&CfRule> = rules.iter().filter(|r| r.schema == CfRuleSchema::X14).collect();
    if x14_rules.is_empty() {
        return None;
    }

    let mut s = String::new();
    s.push_str("<extLst>");
    s.push_str(r#"<ext uri=""#);
    s.push_str(X14_CONDITIONAL_FORMATTING_URI);
    s.push_str(r#"" xmlns:x14=""#);
    s.push_str(X14_NS);
    s.push_str(r#"">"#);
    s.push_str("<x14:conditionalFormattings>");

    for rule in x14_rules {
        s.push_str(r#"<x14:conditionalFormatting xmlns:xm=""#);
        s.push_str(XM_NS);
        s.push_str(r#"">"#);
        s.push_str(&write_x14_cf_rule(rule));
        s.push_str("<xm:sqref>");
        s.push_str(&escape_xml(&sqref_for_ranges(&rule.applies_to)));
        s.push_str("</xm:sqref>");
        s.push_str("</x14:conditionalFormatting>");
    }

    s.push_str("</x14:conditionalFormattings>");
    s.push_str("</ext>");
    s.push_str("</extLst>");
    Some(s)
}

fn write_x14_cf_rule(rule: &CfRule) -> String {
    match &rule.kind {
        CfRuleKind::DataBar(db) => {
            let mut s = String::new();
            s.push_str(r#"<x14:cfRule type="dataBar""#);
            if let Some(id) = &rule.id {
                s.push_str(r#" id=""#);
                s.push_str(&escape_xml(id));
                s.push('"');
            }
            s.push('>');
            s.push_str(&write_x14_data_bar(db));
            s.push_str("</x14:cfRule>");
            s
        }
        // Other x14 rule types are not yet modeled for export; fall back to base schema only.
        _ => String::new(),
    }
}

fn write_x14_data_bar(rule: &DataBarRule) -> String {
    let mut s = String::new();
    s.push_str("<x14:dataBar");
    if let Some(min) = rule.min_length {
        s.push_str(r#" minLength=""#);
        s.push_str(&min.to_string());
        s.push('"');
    }
    if let Some(max) = rule.max_length {
        s.push_str(r#" maxLength=""#);
        s.push_str(&max.to_string());
        s.push('"');
    }
    if let Some(gradient) = rule.gradient {
        s.push_str(r#" gradient=""#);
        s.push_str(if gradient { "1" } else { "0" });
        s.push('"');
    }
    // Excel emits direction; default to leftToRight for now.
    s.push_str(r#" direction="leftToRight">"#);
    s.push_str(&write_cfvo(&rule.min, "x14:cfvo"));
    s.push_str(&write_cfvo(&rule.max, "x14:cfvo"));
    s.push_str("</x14:dataBar>");
    s
}

fn write_cfvo(cfvo: &Cfvo, tag: &str) -> String {
    let mut s = String::new();
    s.push('<');
    s.push_str(tag);
    s.push_str(r#" type=""#);
    s.push_str(cfvo_type_attr(cfvo.type_));
    s.push('"');

    if let Some(val) = cfvo.value.as_deref().filter(|v| !v.is_empty()) {
        s.push_str(r#" val=""#);
        if cfvo.type_ == CfvoType::Formula {
            s.push_str(&escape_xml(&normalize_cf_formula(val)));
        } else {
            s.push_str(&escape_xml(val));
        }
        s.push('"');
    }

    s.push_str("/>");
    s
}

fn normalize_cf_formula(formula: &str) -> String {
    let normalized = formula_model::normalize_formula_text(formula).unwrap_or_default();
    crate::formula_text::add_xlfn_prefixes(&normalized)
}

fn sqref_for_ranges(ranges: &[Range]) -> String {
    if ranges.is_empty() {
        return String::new();
    }
    let mut s = String::new();
    for (idx, r) in ranges.iter().enumerate() {
        if idx > 0 {
            s.push(' ');
        }
        s.push_str(&r.to_string());
    }
    s
}

fn cell_is_operator_attr(op: CellIsOperator) -> &'static str {
    match op {
        CellIsOperator::GreaterThan => "greaterThan",
        CellIsOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        CellIsOperator::LessThan => "lessThan",
        CellIsOperator::LessThanOrEqual => "lessThanOrEqual",
        CellIsOperator::Equal => "equal",
        CellIsOperator::NotEqual => "notEqual",
        CellIsOperator::Between => "between",
        CellIsOperator::NotBetween => "notBetween",
    }
}

fn cfvo_type_attr(t: CfvoType) -> &'static str {
    match t {
        CfvoType::Min => "min",
        CfvoType::Max => "max",
        CfvoType::Number => "num",
        CfvoType::Percent => "percent",
        CfvoType::Percentile => "percentile",
        CfvoType::Formula => "formula",
        CfvoType::AutoMin => "autoMin",
        CfvoType::AutoMax => "autoMax",
    }
}

fn icon_set_attr(set: IconSet) -> &'static str {
    match set {
        IconSet::ThreeArrows => "3Arrows",
        IconSet::ThreeTrafficLights1 => "3TrafficLights1",
        IconSet::ThreeTrafficLights2 => "3TrafficLights2",
        IconSet::ThreeFlags => "3Flags",
        IconSet::ThreeSymbols => "3Symbols",
        IconSet::ThreeSymbols2 => "3Symbols2",
        IconSet::FourArrows => "4Arrows",
        IconSet::FourArrowsGray => "4ArrowsGray",
        IconSet::FiveArrows => "5Arrows",
        IconSet::FiveArrowsGray => "5ArrowsGray",
        IconSet::FiveQuarters => "5Quarters",
    }
}

fn color_rgb(color: Option<Color>) -> Option<String> {
    color.and_then(|c| c.argb()).map(|argb| format!("{:08X}", argb))
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::{CfRuleSchema, CfvoType};

    #[test]
    fn writes_xlfn_prefixes_in_conditional_formatting_formulas() {
        let applies_to = vec![Range::from_a1("A1").unwrap()];

        let expression_rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: applies_to.clone(),
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "SEQUENCE(3)".to_string(),
            },
            dependencies: vec![],
        };

        let databar_rule = CfRule {
            schema: CfRuleSchema::X14,
            id: Some("{D4C5B6A7-0000-0000-0000-000000000000}".to_string()),
            priority: 2,
            applies_to,
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::DataBar(DataBarRule {
                min: Cfvo {
                    type_: CfvoType::Min,
                    value: None,
                },
                max: Cfvo {
                    type_: CfvoType::Formula,
                    value: Some("SEQUENCE(3)".to_string()),
                },
                color: Some(Color::new_argb(0xFF638EC6)),
                min_length: Some(0),
                max_length: Some(100),
                gradient: Some(false),
                negative_fill_color: None,
                axis_color: None,
                direction: None,
            }),
            dependencies: vec![],
        };

        let xml = write_conditional_formatting_xml(&[expression_rule, databar_rule]);

        // <cfRule type="expression"><formula>...</formula></cfRule>
        assert!(
            xml.contains("<formula>_xlfn.SEQUENCE(3)</formula>"),
            "expected _xlfn prefix in <formula>, got:\n{xml}"
        );

        // base cfvo formula attribute
        assert!(
            xml.contains(r#"type="formula" val="_xlfn.SEQUENCE(3)""#),
            "expected _xlfn prefix in base cfvo val, got:\n{xml}"
        );

        // x14 cfvo formula attribute
        assert!(
            xml.contains(r#"<x14:cfvo type="formula" val="_xlfn.SEQUENCE(3)"/>"#),
            "expected _xlfn prefix in x14 cfvo val, got:\n{xml}"
        );
    }
}
