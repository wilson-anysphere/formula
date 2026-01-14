use formula_model::calc_settings::{CalcSettings, CalculationMode, IterativeCalculationSettings};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::package::XlsxPackage;
use crate::xml::workbook_xml_namespaces_from_workbook_start;

#[derive(Debug, thiserror::Error)]
pub enum CalcSettingsError {
    #[error("missing required xlsx part: {0}")]
    MissingPart(&'static str),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
    #[error("utf8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
}

impl XlsxPackage {
    pub fn calc_settings(&self) -> Result<CalcSettings, CalcSettingsError> {
        let workbook_xml = self
            .part("xl/workbook.xml")
            .ok_or(CalcSettingsError::MissingPart("xl/workbook.xml"))?;
        read_calc_settings_from_workbook_xml(workbook_xml)
    }

    pub fn set_calc_settings(&mut self, settings: &CalcSettings) -> Result<(), CalcSettingsError> {
        let workbook_xml = self
            .part("xl/workbook.xml")
            .ok_or(CalcSettingsError::MissingPart("xl/workbook.xml"))?
            .to_vec();
        let updated = write_calc_settings_to_workbook_xml(&workbook_xml, settings)?;
        self.set_part("xl/workbook.xml", updated);
        Ok(())
    }
}

pub fn read_calc_settings_from_workbook_xml(
    workbook_xml: &[u8],
) -> Result<CalcSettings, CalcSettingsError> {
    let mut settings = CalcSettings::default();
    let mut reader = Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            // `workbook.xml` elements may be namespace-prefixed (e.g. `<x:calcPr>`). Match
            // by local name so we can parse SpreadsheetML XML regardless of prefix.
            Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"calcPr" => {
                apply_calc_pr_attributes(&mut reader, &e, &mut settings)?;
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(settings)
}

fn apply_calc_pr_attributes(
    reader: &Reader<&[u8]>,
    e: &BytesStart<'_>,
    settings: &mut CalcSettings,
) -> Result<(), CalcSettingsError> {
    let mut iterative_enabled = settings.iterative.enabled;
    let mut max_iterations = settings.iterative.max_iterations;
    let mut max_change = settings.iterative.max_change;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key = attr.key.as_ref();
        let value = attr.unescape_value()?.to_string();
        match key {
            b"calcMode" => {
                if let Some(mode) = CalculationMode::from_calc_mode_attr(&value) {
                    settings.calculation_mode = mode;
                }
            }
            b"calcOnSave" => settings.calculate_before_save = parse_bool_attr(&value),
            b"fullCalcOnLoad" => settings.full_calc_on_load = parse_bool_attr(&value),
            b"fullPrecision" => settings.full_precision = parse_bool_attr(&value),
            b"iterative" => iterative_enabled = parse_bool_attr(&value),
            b"iterateCount" => {
                if let Ok(n) = value.parse::<u32>() {
                    max_iterations = n;
                }
            }
            b"iterateDelta" => {
                if let Ok(n) = value.parse::<f64>() {
                    max_change = n;
                }
            }
            _ => {}
        }
    }

    settings.iterative = IterativeCalculationSettings {
        enabled: iterative_enabled,
        max_iterations,
        max_change,
    };

    let _ = reader;
    Ok(())
}

pub fn write_calc_settings_to_workbook_xml(
    workbook_xml: &[u8],
    settings: &CalcSettings,
) -> Result<Vec<u8>, CalcSettingsError> {
    let mut reader = Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::new());

    let mut buf = Vec::new();
    let mut wrote_calc_pr = false;
    let mut pending_insert_before_workbook_end = false;
    let mut workbook_ns: Option<crate::xml::WorkbookXmlNamespaces> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"workbook" => {
                pending_insert_before_workbook_end = true;
                workbook_ns.get_or_insert(workbook_xml_namespaces_from_workbook_start(&e)?);
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"workbook" => {
                // Some workbooks may have a self-closing `<workbook/>` root. In that case there
                // is no corresponding `</workbook>` event where we can insert `<calcPr/>`, so we
                // need to expand it into an explicit start/end pair and insert the calc settings.
                let workbook_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                let ns = workbook_xml_namespaces_from_workbook_start(&e)?;
                let calc_pr_tag =
                    crate::xml::prefixed_tag(ns.spreadsheetml_prefix.as_deref(), "calcPr");

                writer.write_event(Event::Start(e.to_owned()))?;
                write_calc_pr_event(
                    &mut writer,
                    settings,
                    CalcPrEventKind::Empty,
                    calc_pr_tag.as_str(),
                )?;
                wrote_calc_pr = true;
                writer.write_event(Event::End(BytesEnd::new(workbook_tag.as_str())))?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"calcPr" => {
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                write_calc_pr_event(&mut writer, settings, CalcPrEventKind::Empty, tag.as_str())?;
                wrote_calc_pr = true;
            }
            Event::Start(e) if e.local_name().as_ref() == b"calcPr" => {
                // Replace the start element, then skip until its end.
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                write_calc_pr_event(&mut writer, settings, CalcPrEventKind::Start, tag.as_str())?;
                wrote_calc_pr = true;

                // Consume inner content (calcPr should be empty in Excel, but be defensive).
                let mut inner_buf = Vec::new();
                loop {
                    match reader.read_event_into(&mut inner_buf)? {
                        Event::End(end) if end.local_name().as_ref() == b"calcPr" => break,
                        Event::Eof => break,
                        _ => {}
                    }
                    inner_buf.clear();
                }
                writer.write_event(Event::End(BytesEnd::new(tag.as_str())))?;
            }
            Event::End(e) if e.local_name().as_ref() == b"workbook" => {
                if pending_insert_before_workbook_end && !wrote_calc_pr {
                    let tag = workbook_ns
                        .as_ref()
                        .map(|ns| crate::xml::prefixed_tag(ns.spreadsheetml_prefix.as_deref(), "calcPr"))
                        .unwrap_or_else(|| "calcPr".to_string());
                    write_calc_pr_event(&mut writer, settings, CalcPrEventKind::Empty, tag.as_str())?;
                    wrote_calc_pr = true;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::Eof => break,
            evt => {
                writer.write_event(evt.to_owned())?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CalcPrEventKind {
    Empty,
    Start,
}

fn write_calc_pr_event(
    writer: &mut Writer<Vec<u8>>,
    settings: &CalcSettings,
    kind: CalcPrEventKind,
    tag: &str,
) -> Result<(), quick_xml::Error> {
    let mut calc_pr = BytesStart::new(tag);
    calc_pr.push_attribute(("calcMode", settings.calculation_mode.as_calc_mode_attr()));
    let calc_on_save = bool_attr(settings.calculate_before_save);
    calc_pr.push_attribute(("calcOnSave", calc_on_save.as_str()));
    let full_calc_on_load = bool_attr(settings.full_calc_on_load);
    calc_pr.push_attribute(("fullCalcOnLoad", full_calc_on_load.as_str()));
    let iterative = bool_attr(settings.iterative.enabled);
    calc_pr.push_attribute(("iterative", iterative.as_str()));
    let iterate_count = settings.iterative.max_iterations.to_string();
    calc_pr.push_attribute(("iterateCount", iterate_count.as_str()));
    let iterate_delta = trim_float(settings.iterative.max_change);
    calc_pr.push_attribute(("iterateDelta", iterate_delta.as_str()));
    let full_precision = bool_attr(settings.full_precision);
    calc_pr.push_attribute(("fullPrecision", full_precision.as_str()));

    match kind {
        CalcPrEventKind::Empty => writer.write_event(Event::Empty(calc_pr))?,
        CalcPrEventKind::Start => writer.write_event(Event::Start(calc_pr))?,
    }
    Ok(())
}

fn parse_bool_attr(value: &str) -> bool {
    matches!(value, "1" | "true" | "TRUE")
}

fn bool_attr(value: bool) -> String {
    if value { "1" } else { "0" }.to_string()
}

fn trim_float(value: f64) -> String {
    // Avoid serializing `-0` for `-0.0` inputs; Excel will treat it as zero but it produces
    // unnecessary diffs/noise.
    if value == 0.0 {
        return "0".to_string();
    }
    let s = format!("{value:.15}");
    let s = s.trim_end_matches('0').trim_end_matches('.');
    if s.is_empty() {
        "0".to_string()
    } else {
        s.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn read_defaults_when_calc_pr_missing() {
        let xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></workbook>"#;
        let settings = read_calc_settings_from_workbook_xml(xml).unwrap();
        assert_eq!(settings, CalcSettings::default());
    }

    #[test]
    fn read_calc_pr_with_prefix_only_workbook_xml() {
        let xml = br#"
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:calcPr
    calcMode="manual"
    calcOnSave="0"
    fullCalcOnLoad="1"
    fullPrecision="0"
    iterative="1"
    iterateCount="42"
    iterateDelta="0.25"
  />
</x:workbook>
"#;
        let settings = read_calc_settings_from_workbook_xml(xml).unwrap();
        assert_eq!(settings.calculation_mode, CalculationMode::Manual);
        assert!(!settings.calculate_before_save);
        assert!(settings.full_calc_on_load);
        assert!(!settings.full_precision);
        assert!(settings.iterative.enabled);
        assert_eq!(settings.iterative.max_iterations, 42);
        assert!((settings.iterative.max_change - 0.25).abs() < 1e-12);
    }

    #[test]
    fn write_calc_pr_into_prefixed_self_closing_workbook_root() {
        let xml = br#"<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
        let updated = write_calc_settings_to_workbook_xml(xml, &CalcSettings::default()).unwrap();
        let updated = std::str::from_utf8(&updated).unwrap();

        assert!(updated.contains("<x:calcPr"));
        assert!(updated.contains("</x:workbook>"));
    }

    #[test]
    fn write_calc_pr_into_default_ns_self_closing_workbook_root() {
        let xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
        let updated = write_calc_settings_to_workbook_xml(xml, &CalcSettings::default()).unwrap();
        let updated = std::str::from_utf8(&updated).unwrap();

        assert!(updated.contains("<calcPr"));
        assert!(updated.contains("</workbook>"));
        assert!(!updated.contains(":calcPr"));
    }

    #[test]
    fn trim_float_serializes_negative_zero_as_zero() {
        assert_eq!(trim_float(-0.0), "0");
        assert_eq!(trim_float(0.0), "0");
    }
}
