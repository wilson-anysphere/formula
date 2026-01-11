use std::collections::BTreeMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};

/// Policy describing how writers should ensure Excel recalculates formulas after edits.
///
/// Excel workbooks may contain both cached `<v>` values and an optional `xl/calcChain.xml`. If we
/// edit formulas without updating the calc chain, Excel can open the file with stale calculation
/// state. The safest approach is to drop the calc chain and request a full calculation on load.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RecalcPolicy {
    /// When a formula-affecting edit occurs, set `<calcPr fullCalcOnLoad="1"/>` in `xl/workbook.xml`.
    pub force_full_calc_on_formula_change: bool,
    /// When a formula-affecting edit occurs, remove `xl/calcChain.xml` and the associated metadata
    /// entries from the package (`[Content_Types].xml` + `xl/_rels/workbook.xml.rels`).
    pub drop_calc_chain_on_formula_change: bool,
    /// When a formula-affecting edit occurs, clear cached `<v>` values for edited formula cells.
    ///
    /// This can help avoid Excel briefly displaying stale values before recalculation on open.
    /// Default: `false`.
    pub clear_cached_values_on_formula_change: bool,
}

impl RecalcPolicy {
    /// Preserve existing calculation metadata (do nothing).
    pub const PRESERVE: Self = Self {
        force_full_calc_on_formula_change: false,
        drop_calc_chain_on_formula_change: false,
        clear_cached_values_on_formula_change: false,
    };
}

impl Default for RecalcPolicy {
    fn default() -> Self {
        Self {
            force_full_calc_on_formula_change: true,
            drop_calc_chain_on_formula_change: true,
            clear_cached_values_on_formula_change: false,
        }
    }
}

#[derive(Debug, thiserror::Error)]
pub(crate) enum RecalcPolicyError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
}

pub(crate) fn apply_recalc_policy_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    policy: RecalcPolicy,
) -> Result<(), RecalcPolicyError> {
    if !policy.force_full_calc_on_formula_change && !policy.drop_calc_chain_on_formula_change {
        return Ok(());
    }

    if policy.force_full_calc_on_formula_change {
        if let Some(workbook_xml) = parts.get("xl/workbook.xml").cloned() {
            let updated = workbook_xml_force_full_calc_on_load(&workbook_xml)?;
            parts.insert("xl/workbook.xml".to_string(), updated);
        }
    }

    if policy.drop_calc_chain_on_formula_change {
        parts.remove("xl/calcChain.xml");

        if let Some(rels_xml) = parts.get("xl/_rels/workbook.xml.rels").cloned() {
            let updated = workbook_rels_remove_calc_chain(&rels_xml)?;
            parts.insert("xl/_rels/workbook.xml.rels".to_string(), updated);
        }

        if let Some(content_types_xml) = parts.get("[Content_Types].xml").cloned() {
            let updated = content_types_remove_calc_chain(&content_types_xml)?;
            parts.insert("[Content_Types].xml".to_string(), updated);
        }
    }

    Ok(())
}

pub(crate) fn workbook_xml_force_full_calc_on_load(
    workbook_xml: &[u8],
) -> Result<Vec<u8>, RecalcPolicyError> {
    let mut reader = Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(workbook_xml.len() + 64));

    let mut buf = Vec::new();
    let mut saw_calc_pr = false;
    let mut in_workbook = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) if e.name().as_ref() == b"workbook" => {
                in_workbook = true;
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if e.name().as_ref() == b"calcPr" => {
                saw_calc_pr = true;
                writer.write_event(Event::Empty(patched_calc_pr(e)?))?;
            }
            Event::Start(ref e) if e.name().as_ref() == b"calcPr" => {
                saw_calc_pr = true;
                writer.write_event(Event::Start(patched_calc_pr(e)?))?;
            }
            Event::End(ref e) if e.name().as_ref() == b"workbook" => {
                if in_workbook && !saw_calc_pr {
                    let mut calc_pr = BytesStart::new("calcPr");
                    calc_pr.push_attribute(("fullCalcOnLoad", "1"));
                    writer.write_event(Event::Empty(calc_pr))?;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::Eof => break,
            other => writer.write_event(other.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn patched_calc_pr(e: &BytesStart<'_>) -> Result<BytesStart<'static>, RecalcPolicyError> {
    let mut calc_pr = BytesStart::new("calcPr");
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"fullCalcOnLoad" {
            continue;
        }
        calc_pr.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
    }
    calc_pr.push_attribute(("fullCalcOnLoad", "1"));
    Ok(calc_pr.into_owned())
}

pub(crate) fn workbook_rels_remove_calc_chain(rels_xml: &[u8]) -> Result<Vec<u8>, RecalcPolicyError> {
    const CALC_CHAIN_REL_TYPE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(rels_xml.len()));

    let mut buf = Vec::new();
    let mut skipping = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if e.name().as_ref() == b"Relationship" => {
                if relationship_is_calc_chain(e, CALC_CHAIN_REL_TYPE)? {
                    skipping = true;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e) if e.name().as_ref() == b"Relationship" => {
                if !relationship_is_calc_chain(e, CALC_CHAIN_REL_TYPE)? {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) if skipping && e.name().as_ref() == b"Relationship" => {
                skipping = false;
            }
            ev if skipping => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn relationship_is_calc_chain(e: &BytesStart<'_>, expected_type: &str) -> Result<bool, RecalcPolicyError> {
    let mut rel_type: Option<String> = None;
    let mut target: Option<String> = None;
    for attr in e.attributes() {
        let attr = attr?;
        let value = attr.unescape_value()?.into_owned();
        match attr.key.as_ref() {
            b"Type" => rel_type = Some(value),
            b"Target" => target = Some(value),
            _ => {}
        }
    }

    if rel_type.as_deref() == Some(expected_type) {
        return Ok(true);
    }

    Ok(target
        .as_deref()
        .is_some_and(|t| t.ends_with("calcChain.xml")))
}

pub(crate) fn content_types_remove_calc_chain(ct_xml: &[u8]) -> Result<Vec<u8>, RecalcPolicyError> {
    let mut reader = Reader::from_reader(ct_xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(ct_xml.len()));

    let mut buf = Vec::new();
    let mut skipping = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if e.name().as_ref() == b"Override" => {
                if override_part_name_is_calc_chain(e)? {
                    skipping = true;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e) if e.name().as_ref() == b"Override" => {
                if !override_part_name_is_calc_chain(e)? {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) if skipping && e.name().as_ref() == b"Override" => {
                skipping = false;
            }
            ev if skipping => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn override_part_name_is_calc_chain(e: &BytesStart<'_>) -> Result<bool, RecalcPolicyError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"PartName" {
            let value = attr.unescape_value()?;
            return Ok(value.as_ref().ends_with("calcChain.xml"));
        }
    }
    Ok(false)
}
