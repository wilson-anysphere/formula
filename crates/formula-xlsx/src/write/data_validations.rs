use formula_model::{
    DataValidationAssignment, DataValidationErrorStyle, DataValidationKind, DataValidationOperator,
};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::XlsxError;

fn bool_attr(value: bool) -> &'static str {
    if value { "1" } else { "0" }
}

fn kind_to_ooxml(kind: DataValidationKind) -> &'static str {
    match kind {
        DataValidationKind::Whole => "whole",
        DataValidationKind::Decimal => "decimal",
        DataValidationKind::List => "list",
        DataValidationKind::Date => "date",
        DataValidationKind::Time => "time",
        DataValidationKind::TextLength => "textLength",
        DataValidationKind::Custom => "custom",
    }
}

fn operator_to_ooxml(op: DataValidationOperator) -> &'static str {
    match op {
        DataValidationOperator::Between => "between",
        DataValidationOperator::NotBetween => "notBetween",
        DataValidationOperator::Equal => "equal",
        DataValidationOperator::NotEqual => "notEqual",
        DataValidationOperator::GreaterThan => "greaterThan",
        DataValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        DataValidationOperator::LessThan => "lessThan",
        DataValidationOperator::LessThanOrEqual => "lessThanOrEqual",
    }
}

fn error_style_to_ooxml(style: DataValidationErrorStyle) -> &'static str {
    match style {
        DataValidationErrorStyle::Stop => "stop",
        DataValidationErrorStyle::Warning => "warning",
        DataValidationErrorStyle::Information => "information",
    }
}

fn insert_before_tag(name: &[u8]) -> bool {
    matches!(
        name,
        // Elements that come after <dataValidations> in the SpreadsheetML schema.
        b"hyperlinks"
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
    )
}

/// Update (or remove) a worksheet `<dataValidations>` block to match `data_validations`.
///
/// If the worksheet already contains `<dataValidations>`, it is replaced. If it does not
/// and `data_validations` is non-empty, the block is inserted before the end of the worksheet
/// (preferably before elements that are required to come after it, e.g. `<hyperlinks>`,
/// `<pageMargins>`, `<tableParts>`, `<extLst>`).
pub(crate) fn update_worksheet_data_validations_xml(
    sheet_xml: &str,
    data_validations: &[DataValidationAssignment],
) -> Result<String, XlsxError> {
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(sheet_xml)?;
    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut skip_depth: usize = 0;
    let mut replaced = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"dataValidations" => {
                replaced = true;
                if !data_validations.is_empty() {
                    write_data_validations_block(
                        &mut writer,
                        data_validations,
                        worksheet_prefix.as_deref(),
                    )?;
                }
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"dataValidations" => {
                replaced = true;
                if !data_validations.is_empty() {
                    write_data_validations_block(
                        &mut writer,
                        data_validations,
                        worksheet_prefix.as_deref(),
                    )?;
                }
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if !replaced
                    && !data_validations.is_empty()
                    && insert_before_tag(e.local_name().as_ref()) =>
            {
                write_data_validations_block(
                    &mut writer,
                    data_validations,
                    worksheet_prefix.as_deref(),
                )?;
                replaced = true;
                writer.write_event(event.to_owned())?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if !replaced && !data_validations.is_empty() {
                    write_data_validations_block(
                        &mut writer,
                        data_validations,
                        worksheet_prefix.as_deref(),
                    )?;
                    replaced = true;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    Ok(String::from_utf8(writer.into_inner())?)
}

fn write_data_validations_block<W: std::io::Write>(
    writer: &mut Writer<W>,
    data_validations: &[DataValidationAssignment],
    prefix: Option<&str>,
) -> Result<(), XlsxError> {
    let data_validations_tag = crate::xml::prefixed_tag(prefix, "dataValidations");
    let data_validation_tag = crate::xml::prefixed_tag(prefix, "dataValidation");
    let formula1_tag = crate::xml::prefixed_tag(prefix, "formula1");
    let formula2_tag = crate::xml::prefixed_tag(prefix, "formula2");

    let count = data_validations.len().to_string();
    let mut start = BytesStart::new(data_validations_tag.as_str());
    start.push_attribute(("count", count.as_str()));
    writer.write_event(Event::Start(start))?;

    for assignment in data_validations {
        let validation = &assignment.validation;

        let mut ranges = assignment.ranges.clone();
        ranges.sort_by_key(|r| (r.start.row, r.start.col, r.end.row, r.end.col));
        let sqref = ranges
            .iter()
            .map(|r| r.to_string())
            .collect::<Vec<_>>()
            .join(" ");

        let mut dv = BytesStart::new(data_validation_tag.as_str());
        dv.push_attribute(("type", kind_to_ooxml(validation.kind)));
        if let Some(op) = validation.operator {
            dv.push_attribute(("operator", operator_to_ooxml(op)));
        }

        dv.push_attribute(("allowBlank", bool_attr(validation.allow_blank)));
        dv.push_attribute(("showInputMessage", bool_attr(validation.show_input_message)));
        dv.push_attribute(("showErrorMessage", bool_attr(validation.show_error_message)));
        dv.push_attribute(("showDropDown", bool_attr(validation.show_drop_down)));

        if let Some(msg) = &validation.input_message {
            if let Some(title) = msg.title.as_deref() {
                dv.push_attribute(("promptTitle", title));
            }
            if let Some(body) = msg.body.as_deref() {
                dv.push_attribute(("prompt", body));
            }
        }

        if let Some(alert) = &validation.error_alert {
            dv.push_attribute(("errorStyle", error_style_to_ooxml(alert.style)));
            if let Some(title) = alert.title.as_deref() {
                dv.push_attribute(("errorTitle", title));
            }
            if let Some(body) = alert.body.as_deref() {
                dv.push_attribute(("error", body));
            }
        }

        dv.push_attribute(("sqref", sqref.as_str()));

        writer.write_event(Event::Start(dv))?;

        // `formula1`/`formula2` are stored in the model without a leading '=' but we accept '='
        // defensively and always strip it when writing. Also restore `_xlfn.` prefixes for
        // forward-compatible functions, mirroring the cell formula and defined-name writers.
        let formula1 = crate::formula_text::add_xlfn_prefixes(super::strip_leading_equals(
            validation.formula1.as_str(),
        ));
        writer.write_event(Event::Start(BytesStart::new(formula1_tag.as_str())))?;
        writer
            .get_mut()
            .write_all(super::escape_text(&formula1).as_bytes())?;
        writer.write_event(Event::End(BytesEnd::new(formula1_tag.as_str())))?;

        if let Some(formula2_raw) = validation.formula2.as_deref() {
            let formula2 = crate::formula_text::add_xlfn_prefixes(super::strip_leading_equals(
                formula2_raw,
            ));
            writer.write_event(Event::Start(BytesStart::new(formula2_tag.as_str())))?;
            writer
                .get_mut()
                .write_all(super::escape_text(&formula2).as_bytes())?;
            writer.write_event(Event::End(BytesEnd::new(formula2_tag.as_str())))?;
        }

        writer.write_event(Event::End(BytesEnd::new(data_validation_tag.as_str())))?;
    }

    writer.write_event(Event::End(BytesEnd::new(data_validations_tag.as_str())))?;
    Ok(())
}
