use std::borrow::Cow;
use std::collections::HashMap;

use formula_model::{
    CellIsOperator, CfRule, CfRuleKind, CfRuleSchema, Cfvo, CfvoType, ColorScaleRule,
    DataBarDirection, DataBarRule, IconSet, IconSetRule, Range, TopBottomKind, UniqueDuplicateRule,
};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::XlsxError;

const X14_CONDITIONAL_FORMATTING_EXT_URI: &str = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
const NS_X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const NS_XM: &str = "http://schemas.microsoft.com/office/excel/2006/main";

fn insert_before_tag(name: &[u8]) -> bool {
    matches!(
        name,
        // Elements that come after <conditionalFormatting> in the SpreadsheetML schema.
        b"dataValidations"
            | b"hyperlinks"
            | b"printOptions"
            | b"pageMargins"
            | b"pageSetup"
            | b"headerFooter"
            | b"rowBreaks"
            | b"colBreaks"
            | b"customProperties"
            | b"cellWatches"
            | b"ignoredErrors"
            | b"smartTags"
            | b"drawing"
            | b"drawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
            // Markup-compatibility blocks: insert known schema elements before these so we don't
            // place SpreadsheetML children after <mc:AlternateContent>.
            | b"AlternateContent"
    )
}

/// Update (or remove) worksheet conditional formatting to match `rules`.
///
/// This performs a streaming rewrite of the worksheet XML:
/// - Removes all existing top-level `<conditionalFormatting>` blocks.
/// - Inserts new `<conditionalFormatting>` blocks rendered from `rules` in the correct schema
///   position (after `<mergeCells>` / `<phoneticPr>`, before `<dataValidations>`, `<hyperlinks>`,
///   `<extLst>`, etc).
/// - Ensures the x14 conditional formatting `extLst` entry is present/updated when needed.
///
/// This function intentionally avoids DOM-parsing the worksheet and preserves unrelated XML
/// subtrees (including `mc:AlternateContent`) by copying events through `quick_xml`.
pub fn update_worksheet_conditional_formatting_xml(
    sheet_xml: &str,
    rules: &[CfRule],
) -> Result<String, XlsxError> {
    // Preserve the original API for callers outside the writer. A fixed seed still guarantees
    // deterministic output given the same input rules.
    update_worksheet_conditional_formatting_xml_with_seed(sheet_xml, rules, 0)
}

/// Update worksheet conditional formatting to match `rules`, generating deterministic ids for
/// x14 rules using `seed`.
///
/// This is identical to [`update_worksheet_conditional_formatting_xml`], but allows callers (e.g.
/// the workbook writer) to incorporate worksheet identity into the generated ids (recommended).
pub fn update_worksheet_conditional_formatting_xml_with_seed(
    sheet_xml: &str,
    rules: &[CfRule],
    seed: u128,
) -> Result<String, XlsxError> {
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(sheet_xml)?;

    // x14 conditional formatting requires `cfRule/@id` so the base rule can link to the x14
    // extension rule. If callers haven't assigned ids yet, generate deterministic ones so writer
    // output is stable across runs (important for tests and reproducible fixtures).
    //
    // Note: ids are only required for `CfRuleSchema::X14`. We intentionally avoid synthesizing ids
    // for Office 2007 rules because Excel typically omits them, and adding them would introduce
    // noisy diffs in patched worksheets.
    let rules: Cow<'_, [CfRule]> = if rules.iter().any(|r| {
        r.schema == CfRuleSchema::X14 && r.id.as_deref().map(|s| s.is_empty()).unwrap_or(true)
    }) {
        let mut owned = rules.to_vec();
        let had_id: Vec<bool> = owned
            .iter()
            .map(|r| r.id.as_deref().filter(|s| !s.is_empty()).is_some())
            .collect();

        super::ensure_rule_ids(&mut owned, seed);

        for (rule, had_id) in owned.iter_mut().zip(had_id) {
            if rule.schema != CfRuleSchema::X14 && !had_id {
                rule.id = None;
            }
        }

        Cow::Owned(owned)
    } else {
        Cow::Borrowed(rules)
    };
    let rules: &[CfRule] = rules.as_ref();

    let needs_base_cf = !rules.is_empty();
    let needs_x14 = rules.iter().any(|r| r.schema == CfRuleSchema::X14);

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(sheet_xml.len() + 256));

    let mut buf = Vec::new();

    let mut in_worksheet = false;
    // Depth inside the worksheet root (root itself excluded).
    let mut ws_depth: usize = 0;

    let mut inserted_cf = false;
    let mut saw_extlst = false;
    let mut in_extlst = false;
    let mut wrote_x14_ext = false;

    // Skip depth for swallowing a top-level `<conditionalFormatting>` subtree.
    let mut skip_cf_depth: usize = 0;
    // Skip depth for swallowing the target `<ext uri="{...}">` subtree inside `<extLst>`.
    let mut skip_ext_depth: usize = 0;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            // Handle a degenerate `<worksheet/>` by expanding it if we need to insert any blocks.
            Event::Empty(ref e) if !in_worksheet && e.local_name().as_ref() == b"worksheet" => {
                if !needs_base_cf && !needs_x14 {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                    break;
                }

                let worksheet_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.write_event(Event::Start(e.to_owned()))?;

                if needs_base_cf {
                    write_conditional_formatting_blocks(&mut writer, rules, worksheet_prefix.as_deref())?;
                }
                if needs_x14 {
                    write_extlst_with_x14(&mut writer, rules, worksheet_prefix.as_deref())?;
                }

                writer.write_event(Event::End(BytesEnd::new(worksheet_tag.as_str())))?;
                break;
            }
            Event::Start(ref e) if !in_worksheet && e.local_name().as_ref() == b"worksheet" => {
                in_worksheet = true;
                ws_depth = 0;
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            // Skip existing top-level conditionalFormatting blocks.
            _ if skip_cf_depth > 0 => {
                match event {
                    Event::Start(_) => skip_cf_depth += 1,
                    Event::End(_) => skip_cf_depth = skip_cf_depth.saturating_sub(1),
                    Event::Empty(_) => {}
                    _ => {}
                }
            }
            // Skip existing target ext blocks inside extLst.
            _ if skip_ext_depth > 0 => {
                match event {
                    Event::Start(_) => skip_ext_depth += 1,
                    Event::End(_) => skip_ext_depth = skip_ext_depth.saturating_sub(1),
                    Event::Empty(_) => {}
                    _ => {}
                }
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if in_worksheet
                    && ws_depth == 0
                    && e.local_name().as_ref() == b"conditionalFormatting" =>
            {
                // Replace the first encountered existing conditionalFormatting block by inserting
                // our new blocks in-place.
                if !inserted_cf && needs_base_cf {
                    write_conditional_formatting_blocks(
                        &mut writer,
                        rules,
                        worksheet_prefix.as_deref(),
                    )?;
                    inserted_cf = true;
                } else if !inserted_cf {
                    // No rules to insert; still mark as inserted so we don't try again later.
                    inserted_cf = true;
                }

                if matches!(event, Event::Start(_)) {
                    skip_cf_depth = 1;
                }
            }
            Event::Start(ref e) if in_worksheet && ws_depth == 0 && e.local_name().as_ref() == b"extLst" => {
                // `extLst` is always after `<conditionalFormatting>` in the schema, so it is a
                // safe insertion point for new conditional formatting blocks if we haven't
                // written them yet.
                if !inserted_cf && needs_base_cf {
                    write_conditional_formatting_blocks(
                        &mut writer,
                        rules,
                        worksheet_prefix.as_deref(),
                    )?;
                    inserted_cf = true;
                }
                saw_extlst = true;
                in_extlst = true;
                wrote_x14_ext = false;
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if in_worksheet && ws_depth == 0 && e.local_name().as_ref() == b"extLst" => {
                if !inserted_cf && needs_base_cf {
                    write_conditional_formatting_blocks(
                        &mut writer,
                        rules,
                        worksheet_prefix.as_deref(),
                    )?;
                    inserted_cf = true;
                }
                saw_extlst = true;
                if needs_x14 {
                    let extlst_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Start(e.to_owned()))?;
                    write_x14_ext_entry(&mut writer, rules, worksheet_prefix.as_deref())?;
                    writer.write_event(Event::End(BytesEnd::new(extlst_tag.as_str())))?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if in_worksheet
                    && ws_depth == 0
                    && !inserted_cf
                    && needs_base_cf
                    && insert_before_tag(e.local_name().as_ref()) =>
            {
                write_conditional_formatting_blocks(
                    &mut writer,
                    rules,
                    worksheet_prefix.as_deref(),
                )?;
                inserted_cf = true;
                writer.write_event(event.to_owned())?;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if in_extlst && ws_depth == 1 && e.local_name().as_ref() == b"ext" =>
            {
                if is_x14_cf_ext_entry(e)? {
                    if needs_x14 && !wrote_x14_ext {
                        write_x14_ext_entry(&mut writer, rules, worksheet_prefix.as_deref())?;
                        wrote_x14_ext = true;
                    }
                    if matches!(event, Event::Start(_)) {
                        skip_ext_depth = 1;
                    }
                } else {
                    writer.write_event(event.to_owned())?;
                }
            }
            Event::End(ref e) if in_extlst && e.local_name().as_ref() == b"extLst" => {
                if needs_x14 && !wrote_x14_ext {
                    write_x14_ext_entry(&mut writer, rules, worksheet_prefix.as_deref())?;
                    wrote_x14_ext = true;
                }
                in_extlst = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::End(ref e) if in_worksheet && e.local_name().as_ref() == b"worksheet" => {
                // If we never inserted CF blocks and need them, insert before </worksheet>.
                if !inserted_cf && needs_base_cf {
                    write_conditional_formatting_blocks(
                        &mut writer,
                        rules,
                        worksheet_prefix.as_deref(),
                    )?;
                    inserted_cf = true;
                }

                // If no extLst exists and we have x14 rules, append a new extLst at the end.
                if needs_x14 && !saw_extlst {
                    write_extlst_with_x14(&mut writer, rules, worksheet_prefix.as_deref())?;
                }

                in_worksheet = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }

        // Update depth tracking after handling the event, based on the input structure.
        if in_worksheet {
            match event {
                Event::Start(ref e) if e.local_name().as_ref() != b"worksheet" => {
                    ws_depth = ws_depth.saturating_add(1);
                }
                Event::End(ref e) if e.local_name().as_ref() != b"worksheet" => {
                    ws_depth = ws_depth.saturating_sub(1);
                }
                _ => {}
            }
        }

        buf.clear();
    }

    Ok(String::from_utf8(writer.into_inner())?)
}

fn is_x14_cf_ext_entry(e: &BytesStart<'_>) -> Result<bool, XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if attr.key.as_ref() == b"uri" && attr.value.as_ref() == X14_CONDITIONAL_FORMATTING_EXT_URI.as_bytes() {
            return Ok(true);
        }
    }
    Ok(false)
}

fn write_extlst_with_x14<W: std::io::Write>(
    writer: &mut Writer<W>,
    rules: &[CfRule],
    prefix: Option<&str>,
) -> Result<(), XlsxError> {
    let extlst_tag = crate::xml::prefixed_tag(prefix, "extLst");
    writer.write_event(Event::Start(BytesStart::new(extlst_tag.as_str())))?;
    write_x14_ext_entry(writer, rules, prefix)?;
    writer.write_event(Event::End(BytesEnd::new(extlst_tag.as_str())))?;
    Ok(())
}

fn write_x14_ext_entry<W: std::io::Write>(
    writer: &mut Writer<W>,
    rules: &[CfRule],
    prefix: Option<&str>,
) -> Result<(), XlsxError> {
    let ext_tag = crate::xml::prefixed_tag(prefix, "ext");
    let mut start = BytesStart::new(ext_tag.as_str());
    start.push_attribute(("uri", X14_CONDITIONAL_FORMATTING_EXT_URI));
    start.push_attribute(("xmlns:x14", NS_X14));
    writer.write_event(Event::Start(start))?;

    writer.write_event(Event::Start(BytesStart::new("x14:conditionalFormattings")))?;

    // Group x14 rules by sqref.
    let mut groups: Vec<(String, Vec<&CfRule>)> = Vec::new();
    let mut idx_by_sqref: HashMap<String, usize> = HashMap::new();
    for rule in rules.iter().filter(|r| r.schema == CfRuleSchema::X14) {
        let sqref = format_sqref(&rule.applies_to);
        let idx = *idx_by_sqref.entry(sqref.clone()).or_insert_with(|| {
            groups.push((sqref.clone(), Vec::new()));
            groups.len() - 1
        });
        groups[idx].1.push(rule);
    }

    for (sqref, group_rules) in groups {
        let mut cf_start = BytesStart::new("x14:conditionalFormatting");
        cf_start.push_attribute(("xmlns:xm", NS_XM));
        writer.write_event(Event::Start(cf_start))?;

        for rule in group_rules {
            write_x14_cf_rule(writer, rule)?;
        }

        // `<xm:sqref>` comes last in the block.
        writer.write_event(Event::Start(BytesStart::new("xm:sqref")))?;
        writer.write_event(Event::Text(BytesText::new(sqref.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("xm:sqref")))?;

        writer.write_event(Event::End(BytesEnd::new("x14:conditionalFormatting")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("x14:conditionalFormattings")))?;
    writer.write_event(Event::End(BytesEnd::new(ext_tag.as_str())))?;
    Ok(())
}

fn write_x14_cf_rule<W: std::io::Write>(writer: &mut Writer<W>, rule: &CfRule) -> Result<(), XlsxError> {
    match &rule.kind {
        CfRuleKind::DataBar(db) => write_x14_data_bar_rule(writer, rule, db),
        // Best-effort: non-dataBar rules currently have no x14 payload in `formula_model`.
        _ => Ok(()),
    }
}

fn write_x14_data_bar_rule<W: std::io::Write>(
    writer: &mut Writer<W>,
    rule: &CfRule,
    db: &DataBarRule,
) -> Result<(), XlsxError> {
    let mut start = BytesStart::new("x14:cfRule");
    start.push_attribute(("type", "dataBar"));
    if let Some(id) = rule.id.as_deref() {
        start.push_attribute(("id", id));
    }
    writer.write_event(Event::Start(start))?;

    let mut db_start = BytesStart::new("x14:dataBar");
    if let Some(min) = db.min_length {
        db_start.push_attribute(("minLength", min.to_string().as_str()));
    }
    if let Some(max) = db.max_length {
        db_start.push_attribute(("maxLength", max.to_string().as_str()));
    }
    if let Some(gradient) = db.gradient {
        db_start.push_attribute(("gradient", if gradient { "1" } else { "0" }));
    }
    // Excel emits this even when it matches the default. Include it for compatibility.
    let direction = db.direction.unwrap_or(DataBarDirection::LeftToRight);
    db_start.push_attribute(("direction", data_bar_direction_to_ooxml(direction)));

    writer.write_event(Event::Start(db_start))?;

    write_cfvo(writer, "x14:cfvo", &db.min)?;
    write_cfvo(writer, "x14:cfvo", &db.max)?;

    // Defaults match Excel's typical output.
    let negative_fill_color = db
        .negative_fill_color
        .unwrap_or(formula_model::Color::new_argb(0xFFFF0000));
    let axis_color = db
        .axis_color
        .unwrap_or(formula_model::Color::new_argb(0xFF000000));
    let negative_rgb = format!("{:08X}", negative_fill_color.argb().unwrap_or(0));
    let axis_rgb = format!("{:08X}", axis_color.argb().unwrap_or(0));

    let mut neg = BytesStart::new("x14:negativeFillColor");
    neg.push_attribute(("rgb", negative_rgb.as_str()));
    writer.write_event(Event::Empty(neg))?;
    let mut axis = BytesStart::new("x14:axisColor");
    axis.push_attribute(("rgb", axis_rgb.as_str()));
    writer.write_event(Event::Empty(axis))?;

    writer.write_event(Event::End(BytesEnd::new("x14:dataBar")))?;
    writer.write_event(Event::End(BytesEnd::new("x14:cfRule")))?;
    Ok(())
}

fn write_conditional_formatting_blocks<W: std::io::Write>(
    writer: &mut Writer<W>,
    rules: &[CfRule],
    prefix: Option<&str>,
) -> Result<(), XlsxError> {
    if rules.is_empty() {
        return Ok(());
    }

    let cf_tag = crate::xml::prefixed_tag(prefix, "conditionalFormatting");

    let mut groups: Vec<(String, Vec<&CfRule>)> = Vec::new();
    let mut idx_by_sqref: HashMap<String, usize> = HashMap::new();
    for rule in rules {
        let sqref = format_sqref(&rule.applies_to);
        let idx = *idx_by_sqref.entry(sqref.clone()).or_insert_with(|| {
            groups.push((sqref.clone(), Vec::new()));
            groups.len() - 1
        });
        groups[idx].1.push(rule);
    }

    for (sqref, group_rules) in groups {
        let mut start = BytesStart::new(cf_tag.as_str());
        start.push_attribute(("sqref", sqref.as_str()));
        writer.write_event(Event::Start(start))?;
        for rule in group_rules {
            write_cf_rule(writer, rule, prefix)?;
        }
        writer.write_event(Event::End(BytesEnd::new(cf_tag.as_str())))?;
    }

    Ok(())
}

fn write_cf_rule<W: std::io::Write>(
    writer: &mut Writer<W>,
    rule: &CfRule,
    prefix: Option<&str>,
) -> Result<(), XlsxError> {
    let cf_rule_tag = crate::xml::prefixed_tag(prefix, "cfRule");

    match &rule.kind {
        CfRuleKind::CellIs { operator, formulas } => {
            let mut start = BytesStart::new(cf_rule_tag.as_str());
            start.push_attribute(("type", "cellIs"));
            start.push_attribute(("priority", rule.priority.to_string().as_str()));
            start.push_attribute(("operator", cell_is_operator_to_ooxml(*operator)));
            if let Some(dxf) = rule.dxf_id {
                start.push_attribute(("dxfId", dxf.to_string().as_str()));
            }
            if rule.stop_if_true {
                start.push_attribute(("stopIfTrue", "1"));
            }
            if let Some(id) = rule.id.as_deref() {
                start.push_attribute(("id", id));
            }
            writer.write_event(Event::Start(start))?;

            let formula_tag = crate::xml::prefixed_tag(prefix, "formula");
            for formula in formulas {
                writer.write_event(Event::Start(BytesStart::new(formula_tag.as_str())))?;
                writer.write_event(Event::Text(BytesText::new(formula.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new(formula_tag.as_str())))?;
            }
            writer.write_event(Event::End(BytesEnd::new(cf_rule_tag.as_str())))?;
        }
        CfRuleKind::Expression { formula } => {
            let mut start = BytesStart::new(cf_rule_tag.as_str());
            start.push_attribute(("type", "expression"));
            start.push_attribute(("priority", rule.priority.to_string().as_str()));
            if let Some(dxf) = rule.dxf_id {
                start.push_attribute(("dxfId", dxf.to_string().as_str()));
            }
            if rule.stop_if_true {
                start.push_attribute(("stopIfTrue", "1"));
            }
            if let Some(id) = rule.id.as_deref() {
                start.push_attribute(("id", id));
            }
            writer.write_event(Event::Start(start))?;

            let formula_tag = crate::xml::prefixed_tag(prefix, "formula");
            writer.write_event(Event::Start(BytesStart::new(formula_tag.as_str())))?;
            writer.write_event(Event::Text(BytesText::new(formula.as_str())))?;
            writer.write_event(Event::End(BytesEnd::new(formula_tag.as_str())))?;
            writer.write_event(Event::End(BytesEnd::new(cf_rule_tag.as_str())))?;
        }
        CfRuleKind::DataBar(db) => {
            let mut start = BytesStart::new(cf_rule_tag.as_str());
            start.push_attribute(("type", "dataBar"));
            start.push_attribute(("priority", rule.priority.to_string().as_str()));
            if let Some(dxf) = rule.dxf_id {
                start.push_attribute(("dxfId", dxf.to_string().as_str()));
            }
            if rule.stop_if_true {
                start.push_attribute(("stopIfTrue", "1"));
            }
            if let Some(id) = rule.id.as_deref() {
                start.push_attribute(("id", id));
            }
            writer.write_event(Event::Start(start))?;

            let data_bar_tag = crate::xml::prefixed_tag(prefix, "dataBar");
            writer.write_event(Event::Start(BytesStart::new(data_bar_tag.as_str())))?;
            write_cfvo(writer, crate::xml::prefixed_tag(prefix, "cfvo").as_str(), &db.min)?;
            write_cfvo(writer, crate::xml::prefixed_tag(prefix, "cfvo").as_str(), &db.max)?;

            if let Some(color) = db.color {
                let color_tag = crate::xml::prefixed_tag(prefix, "color");
                let mut color_el = BytesStart::new(color_tag.as_str());
                color_el.push_attribute(("rgb", format!("{:08X}", color.argb().unwrap_or(0)).as_str()));
                writer.write_event(Event::Empty(color_el))?;
            }
            writer.write_event(Event::End(BytesEnd::new(data_bar_tag.as_str())))?;
            writer.write_event(Event::End(BytesEnd::new(cf_rule_tag.as_str())))?;
        }
        CfRuleKind::ColorScale(ColorScaleRule { cfvos, colors }) => {
            let mut start = BytesStart::new(cf_rule_tag.as_str());
            start.push_attribute(("type", "colorScale"));
            start.push_attribute(("priority", rule.priority.to_string().as_str()));
            if let Some(dxf) = rule.dxf_id {
                start.push_attribute(("dxfId", dxf.to_string().as_str()));
            }
            if rule.stop_if_true {
                start.push_attribute(("stopIfTrue", "1"));
            }
            if let Some(id) = rule.id.as_deref() {
                start.push_attribute(("id", id));
            }
            writer.write_event(Event::Start(start))?;

            let color_scale_tag = crate::xml::prefixed_tag(prefix, "colorScale");
            writer.write_event(Event::Start(BytesStart::new(color_scale_tag.as_str())))?;
            let cfvo_tag = crate::xml::prefixed_tag(prefix, "cfvo");
            for cfvo in cfvos {
                write_cfvo(writer, cfvo_tag.as_str(), cfvo)?;
            }
            let color_tag = crate::xml::prefixed_tag(prefix, "color");
            for color in colors {
                let mut el = BytesStart::new(color_tag.as_str());
                el.push_attribute(("rgb", format!("{:08X}", color.argb().unwrap_or(0)).as_str()));
                writer.write_event(Event::Empty(el))?;
            }
            writer.write_event(Event::End(BytesEnd::new(color_scale_tag.as_str())))?;
            writer.write_event(Event::End(BytesEnd::new(cf_rule_tag.as_str())))?;
        }
        CfRuleKind::IconSet(IconSetRule {
            set,
            cfvos,
            show_value,
            reverse,
        }) => {
            let mut start = BytesStart::new(cf_rule_tag.as_str());
            start.push_attribute(("type", "iconSet"));
            start.push_attribute(("priority", rule.priority.to_string().as_str()));
            if let Some(dxf) = rule.dxf_id {
                start.push_attribute(("dxfId", dxf.to_string().as_str()));
            }
            if rule.stop_if_true {
                start.push_attribute(("stopIfTrue", "1"));
            }
            if let Some(id) = rule.id.as_deref() {
                start.push_attribute(("id", id));
            }
            writer.write_event(Event::Start(start))?;

            let icon_set_tag = crate::xml::prefixed_tag(prefix, "iconSet");
            let mut icon_set_start = BytesStart::new(icon_set_tag.as_str());
            icon_set_start.push_attribute(("iconSet", icon_set_to_ooxml(*set)));
            if !*show_value {
                icon_set_start.push_attribute(("showValue", "0"));
            }
            if *reverse {
                icon_set_start.push_attribute(("reverse", "1"));
            }
            writer.write_event(Event::Start(icon_set_start))?;

            let cfvo_tag = crate::xml::prefixed_tag(prefix, "cfvo");
            for cfvo in cfvos {
                write_cfvo(writer, cfvo_tag.as_str(), cfvo)?;
            }

            writer.write_event(Event::End(BytesEnd::new(icon_set_tag.as_str())))?;
            writer.write_event(Event::End(BytesEnd::new(cf_rule_tag.as_str())))?;
        }
        CfRuleKind::TopBottom(rule_tb) => {
            let mut start = BytesStart::new(cf_rule_tag.as_str());
            start.push_attribute(("type", "top10"));
            start.push_attribute(("priority", rule.priority.to_string().as_str()));
            start.push_attribute(("rank", rule_tb.rank.to_string().as_str()));
            if rule_tb.percent {
                start.push_attribute(("percent", "1"));
            }
            if matches!(rule_tb.kind, TopBottomKind::Bottom) {
                start.push_attribute(("bottom", "1"));
            }
            if let Some(dxf) = rule.dxf_id {
                start.push_attribute(("dxfId", dxf.to_string().as_str()));
            }
            if rule.stop_if_true {
                start.push_attribute(("stopIfTrue", "1"));
            }
            if let Some(id) = rule.id.as_deref() {
                start.push_attribute(("id", id));
            }
            writer.write_event(Event::Empty(start))?;
        }
        CfRuleKind::UniqueDuplicate(UniqueDuplicateRule { unique }) => {
            let mut start = BytesStart::new(cf_rule_tag.as_str());
            start.push_attribute(("type", if *unique { "uniqueValues" } else { "duplicateValues" }));
            start.push_attribute(("priority", rule.priority.to_string().as_str()));
            if let Some(dxf) = rule.dxf_id {
                start.push_attribute(("dxfId", dxf.to_string().as_str()));
            }
            if rule.stop_if_true {
                start.push_attribute(("stopIfTrue", "1"));
            }
            if let Some(id) = rule.id.as_deref() {
                start.push_attribute(("id", id));
            }
            writer.write_event(Event::Empty(start))?;
        }
        CfRuleKind::Unsupported { raw_xml, .. } => {
            // Best-effort: re-emit the stored cfRule XML fragment, if it parses.
            //
            // Note: this assumes the fragment uses the correct element prefixing
            // for the target worksheet.
            let mut frag_reader = Reader::from_str(raw_xml);
            frag_reader.config_mut().trim_text(false);
            let mut frag_buf = Vec::new();
            loop {
                match frag_reader.read_event_into(&mut frag_buf)? {
                    Event::Eof => break,
                    ev => writer.write_event(ev.into_owned())?,
                }
                frag_buf.clear();
            }
        }
    }

    Ok(())
}

fn write_cfvo<W: std::io::Write>(writer: &mut Writer<W>, tag: &str, cfvo: &Cfvo) -> Result<(), XlsxError> {
    let mut el = BytesStart::new(tag);
    el.push_attribute(("type", cfvo_type_to_ooxml(cfvo.type_)));
    if let Some(val) = cfvo.value.as_deref() {
        el.push_attribute(("val", val));
    }
    writer.write_event(Event::Empty(el))?;
    Ok(())
}

fn format_sqref(ranges: &[Range]) -> String {
    ranges
        .iter()
        .map(|r| r.to_string())
        .collect::<Vec<_>>()
        .join(" ")
}

fn cell_is_operator_to_ooxml(op: CellIsOperator) -> &'static str {
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

fn cfvo_type_to_ooxml(type_: CfvoType) -> &'static str {
    match type_ {
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

fn icon_set_to_ooxml(set: IconSet) -> &'static str {
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

fn data_bar_direction_to_ooxml(direction: DataBarDirection) -> &'static str {
    match direction {
        DataBarDirection::LeftToRight => "leftToRight",
        DataBarDirection::RightToLeft => "rightToLeft",
        DataBarDirection::Context => "context",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::{parse_range_a1, Color};
    use roxmltree::Document;
    use std::io::{Cursor, Read};
    use zip::ZipArchive;

    fn expr_rule(range: &str, formula: &str) -> CfRule {
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![parse_range_a1(range).unwrap()],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: formula.to_string(),
            },
            dependencies: vec![],
        }
    }

    fn x14_data_bar_rule(range: &str) -> CfRule {
        CfRule {
            schema: CfRuleSchema::X14,
            id: Some("{A1B2C3D4-E5F6-4711-8899-AABBCCDDEEFF}".to_string()),
            priority: 1,
            applies_to: vec![parse_range_a1(range).unwrap()],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::DataBar(DataBarRule {
                min: Cfvo {
                    type_: CfvoType::AutoMin,
                    value: None,
                },
                max: Cfvo {
                    type_: CfvoType::AutoMax,
                    value: None,
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
        }
    }

    fn x14_data_bar_rule_missing_id(range: &str) -> CfRule {
        let mut rule = x14_data_bar_rule(range);
        rule.id = None;
        rule
    }

    fn is_excel_guid(s: &str) -> bool {
        if s.len() != 38 {
            return false;
        }
        let bytes = s.as_bytes();
        if bytes[0] != b'{' || bytes[37] != b'}' {
            return false;
        }
        if bytes[9] != b'-' || bytes[14] != b'-' || bytes[19] != b'-' || bytes[24] != b'-' {
            return false;
        }
        for (i, &b) in bytes.iter().enumerate() {
            if matches!(i, 0 | 9 | 14 | 19 | 24 | 37) {
                continue;
            }
            let is_hex_upper = (b'0'..=b'9').contains(&b) || (b'A'..=b'F').contains(&b);
            if !is_hex_upper {
                return false;
            }
        }
        true
    }

    #[test]
    fn x14_data_bar_writer_uses_model_direction_and_colors() {
        let rule = CfRule {
            schema: CfRuleSchema::X14,
            id: Some("{11111111-2222-3333-4444-555555555555}".to_string()),
            priority: 1,
            applies_to: vec![parse_range_a1("A1:A1").unwrap()],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::DataBar(DataBarRule {
                min: Cfvo {
                    type_: CfvoType::AutoMin,
                    value: None,
                },
                max: Cfvo {
                    type_: CfvoType::AutoMax,
                    value: None,
                },
                color: Some(Color::new_argb(0xFF638EC6)),
                min_length: Some(0),
                max_length: Some(100),
                gradient: Some(false),
                negative_fill_color: Some(Color::new_argb(0xFF00FF00)),
                axis_color: Some(Color::new_argb(0xFF112233)),
                direction: Some(DataBarDirection::RightToLeft),
            }),
            dependencies: vec![],
        };

        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let updated = update_worksheet_conditional_formatting_xml(xml, &[rule]).unwrap();

        assert!(
            updated.contains(r#"direction="rightToLeft""#),
            "expected direction to roundtrip, got:\n{updated}"
        );
        assert!(
            updated.contains(r#"negativeFillColor rgb="FF00FF00""#),
            "expected negative fill color to roundtrip, got:\n{updated}"
        );
        assert!(
            updated.contains(r#"axisColor rgb="FF112233""#),
            "expected axis color to roundtrip, got:\n{updated}"
        );
    }

    #[test]
    fn inserts_before_data_validations_when_missing() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><dataValidations/></worksheet>"#;
        let rules = vec![expr_rule("A1:A1", "A1>0")];
        let updated = update_worksheet_conditional_formatting_xml(xml, &rules).unwrap();
        let cf_pos = updated.find("<conditionalFormatting").expect("inserted cf");
        let dv_pos = updated.find("<dataValidations").expect("dataValidations exists");
        assert!(cf_pos < dv_pos, "expected CF before dataValidations, got:\n{updated}");
    }

    #[test]
    fn replaces_and_removes_existing_conditional_formatting() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><conditionalFormatting sqref="A1:A1"><cfRule type="expression" priority="1"><formula>OLD</formula></cfRule></conditionalFormatting><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#;
        let rules = vec![expr_rule("B1:B1", "NEW")];
        let updated = update_worksheet_conditional_formatting_xml(xml, &rules).unwrap();
        assert!(!updated.contains("OLD"));
        assert!(updated.contains("NEW"));
        assert!(updated.contains(r#"sqref="B1""#));

        let removed = update_worksheet_conditional_formatting_xml(xml, &[]).unwrap();
        assert!(
            !removed.contains("<conditionalFormatting"),
            "expected CF blocks removed, got:\n{removed}"
        );
    }

    #[test]
    fn x14_extlst_rewrite_preserves_other_ext_entries() {
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><extLst><ext uri="{{OTHER}}"><foo/></ext><ext uri="{x14_uri}" xmlns:x14="{x14_ns}"><x14:conditionalFormattings/></ext></extLst></worksheet>"#,
            x14_uri = X14_CONDITIONAL_FORMATTING_EXT_URI,
            x14_ns = NS_X14
        );

        let rules = vec![x14_data_bar_rule("B1:B3")];
        let updated = update_worksheet_conditional_formatting_xml(&xml, &rules).unwrap();
        assert!(updated.contains(r#"uri="{OTHER}""#));
        assert!(updated.contains("<foo"));
        assert!(updated.contains(X14_CONDITIONAL_FORMATTING_EXT_URI));
        assert!(updated.contains("x14:conditionalFormattings"));

        // Clearing x14 rules should remove just the targeted ext entry, but keep other ext.
        let rules = vec![expr_rule("A1:A1", "A1>0")];
        let cleared = update_worksheet_conditional_formatting_xml(&xml, &rules).unwrap();
        assert!(cleared.contains(r#"uri="{OTHER}""#));
        assert!(cleared.contains("<foo"));
        assert!(
            !cleared.contains(X14_CONDITIONAL_FORMATTING_EXT_URI),
            "expected x14 ext entry removed, got:\n{cleared}"
        );
    }

    #[test]
    fn preserves_mc_alternate_content_blocks() {
        let bytes = include_bytes!("../../tests/fixtures/rt_mc.xlsx");
        let sheet_xml = zip_part(bytes, "xl/worksheets/sheet1.xml");
        let sheet_xml = std::str::from_utf8(&sheet_xml).unwrap();

        let rules = vec![expr_rule("A1:A1", "A1>0")];
        let updated = update_worksheet_conditional_formatting_xml(sheet_xml, &rules).unwrap();
        assert!(updated.contains("mc:AlternateContent"));
        assert!(updated.contains("x14ac:someFutureFeature"));

        // Sanity check: the updated XML is still parseable and retains the MC element.
        let doc = Document::parse(&updated).expect("valid xml");
        let main_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        assert!(doc
            .descendants()
            .any(|n| n.is_element() && n.tag_name().name() == "AlternateContent" && n.tag_name().namespace() == Some(main_ns)));
    }

    #[test]
    fn preserves_worksheet_prefixing() {
        let xml = r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/><x:dataValidations/></x:worksheet>"#;
        let rules = vec![expr_rule("A1:A1", "A1>0")];
        let updated = update_worksheet_conditional_formatting_xml(xml, &rules).unwrap();
        assert!(updated.contains("<x:conditionalFormatting"));
        let cf_pos = updated.find("<x:conditionalFormatting").unwrap();
        let dv_pos = updated.find("<x:dataValidations").unwrap();
        assert!(cf_pos < dv_pos);
    }

    #[test]
    fn x14_data_bar_ext_preserves_direction_and_colors() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let mut rule = x14_data_bar_rule("B1:B3");
        if let CfRuleKind::DataBar(ref mut db) = rule.kind {
            db.direction = Some(DataBarDirection::RightToLeft);
            db.negative_fill_color = Some(Color::new_argb(0xFF112233));
            db.axis_color = Some(Color::new_argb(0xFF445566));
        } else {
            panic!("expected DataBar rule");
        }

        let updated = update_worksheet_conditional_formatting_xml(xml, &[rule]).unwrap();
        let doc = Document::parse(&updated).expect("valid xml");
        let data_bar = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "dataBar" && n.tag_name().namespace() == Some(NS_X14))
            .expect("x14:dataBar should exist");

        assert_eq!(data_bar.attribute("direction"), Some("rightToLeft"));

        let negative_fill = data_bar
            .children()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "negativeFillColor"
                    && n.tag_name().namespace() == Some(NS_X14)
            })
            .expect("x14:negativeFillColor");
        assert_eq!(negative_fill.attribute("rgb"), Some("FF112233"));

        let axis_color = data_bar
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "axisColor" && n.tag_name().namespace() == Some(NS_X14))
            .expect("x14:axisColor");
        assert_eq!(axis_color.attribute("rgb"), Some("FF445566"));
    }

    #[test]
    fn x14_rules_without_ids_get_deterministic_generated_ids_for_linking() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let rules = vec![x14_data_bar_rule_missing_id("B1:B3")];

        let updated1 = update_worksheet_conditional_formatting_xml(xml, &rules).unwrap();
        let updated2 = update_worksheet_conditional_formatting_xml(xml, &rules).unwrap();
        assert_eq!(updated1, updated2, "output must be deterministic across runs");

        let doc = Document::parse(&updated1).expect("valid xml");
        let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        let base_rule = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "cfRule" && n.tag_name().namespace() == Some(main_ns))
            .expect("base cfRule should exist");
        let x14_rule = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "cfRule" && n.tag_name().namespace() == Some(NS_X14))
            .expect("x14 cfRule should exist");

        let base_id = base_rule.attribute("id").expect("base cfRule should have id");
        let x14_id = x14_rule.attribute("id").expect("x14 cfRule should have id");

        assert_eq!(base_id, x14_id, "base and x14 ids must match for linking");
        assert!(is_excel_guid(base_id), "expected Excel GUID format, got {base_id}");
    }

    #[test]
    fn id_generation_does_not_add_ids_to_office2007_rules() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let rules = vec![expr_rule("A1:A1", "A1>0"), x14_data_bar_rule_missing_id("B1:B3")];
        let updated = update_worksheet_conditional_formatting_xml(xml, &rules).unwrap();

        let doc = Document::parse(&updated).expect("valid xml");
        let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        let expr_cf = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "cfRule"
                    && n.tag_name().namespace() == Some(main_ns)
                    && n.attribute("type") == Some("expression")
            })
            .expect("expression cfRule exists");
        assert_eq!(
            expr_cf.attribute("id"),
            None,
            "office2007 expression rules should not have synthesized ids"
        );
    }

    fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
        let cursor = Cursor::new(zip_bytes);
        let mut archive = ZipArchive::new(cursor).expect("open zip");
        let mut file = archive.by_name(name).expect("part exists");
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read part");
        buf
    }
}
