use formula_model::calc_settings::{CalcSettings, CalculationMode, IterativeCalculationSettings};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::package::XlsxPackage;

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
            Event::Empty(e) | Event::Start(e) if e.name().as_ref() == b"calcPr" => {
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

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == b"workbook" => {
                pending_insert_before_workbook_end = true;
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(e) if e.name().as_ref() == b"calcPr" => {
                write_calc_pr_event(&mut writer, settings, CalcPrEventKind::Empty)?;
                wrote_calc_pr = true;
            }
            Event::Start(e) if e.name().as_ref() == b"calcPr" => {
                // Replace the start element, then skip until its end.
                write_calc_pr_event(&mut writer, settings, CalcPrEventKind::Start)?;
                wrote_calc_pr = true;

                // Consume inner content (calcPr should be empty in Excel, but be defensive).
                let mut inner_buf = Vec::new();
                loop {
                    match reader.read_event_into(&mut inner_buf)? {
                        Event::End(end) if end.name().as_ref() == b"calcPr" => break,
                        Event::Eof => break,
                        _ => {}
                    }
                    inner_buf.clear();
                }
                writer.write_event(Event::End(BytesEnd::new("calcPr")))?;
            }
            Event::End(e) if e.name().as_ref() == b"workbook" => {
                if pending_insert_before_workbook_end && !wrote_calc_pr {
                    write_calc_pr_event(&mut writer, settings, CalcPrEventKind::Empty)?;
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
) -> Result<(), quick_xml::Error> {
    let mut calc_pr = BytesStart::new("calcPr");
    calc_pr.push_attribute(("calcMode", settings.calculation_mode.as_calc_mode_attr()));
    let calc_on_save = bool_attr(settings.calculate_before_save);
    calc_pr.push_attribute(("calcOnSave", calc_on_save.as_str()));
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
}
