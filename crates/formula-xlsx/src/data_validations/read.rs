use formula_model::{
    parse_sqref, DataValidation, DataValidationErrorAlert, DataValidationErrorStyle,
    DataValidationInputMessage, DataValidationKind, DataValidationOperator, Range,
};
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::XlsxError;

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct ParsedDataValidation {
    pub ranges: Vec<Range>,
    pub validation: DataValidation,
}

fn parse_xml_bool(val: &str) -> bool {
    val == "1" || val.eq_ignore_ascii_case("true")
}

fn strip_leading_equals(formula: &str) -> &str {
    let trimmed = formula.trim();
    trimmed.strip_prefix('=').unwrap_or(trimmed)
}

fn parse_kind(val: &str) -> Option<DataValidationKind> {
    match val {
        "whole" => Some(DataValidationKind::Whole),
        "decimal" => Some(DataValidationKind::Decimal),
        "list" => Some(DataValidationKind::List),
        "date" => Some(DataValidationKind::Date),
        "time" => Some(DataValidationKind::Time),
        "textLength" => Some(DataValidationKind::TextLength),
        "custom" => Some(DataValidationKind::Custom),
        // Some producers use `none` for a disabled validation; ignore it.
        "none" => None,
        _ => None,
    }
}

fn parse_operator(val: &str) -> Option<DataValidationOperator> {
    match val {
        "between" => Some(DataValidationOperator::Between),
        "notBetween" => Some(DataValidationOperator::NotBetween),
        "equal" => Some(DataValidationOperator::Equal),
        "notEqual" => Some(DataValidationOperator::NotEqual),
        "greaterThan" => Some(DataValidationOperator::GreaterThan),
        "greaterThanOrEqual" => Some(DataValidationOperator::GreaterThanOrEqual),
        "lessThan" => Some(DataValidationOperator::LessThan),
        "lessThanOrEqual" => Some(DataValidationOperator::LessThanOrEqual),
        _ => None,
    }
}

fn parse_error_style(val: &str) -> Option<DataValidationErrorStyle> {
    match val {
        "stop" => Some(DataValidationErrorStyle::Stop),
        "warning" => Some(DataValidationErrorStyle::Warning),
        "information" => Some(DataValidationErrorStyle::Information),
        _ => None,
    }
}

pub(crate) fn read_data_validations_from_worksheet_xml(
    xml: &str,
) -> Result<Vec<ParsedDataValidation>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out: Vec<ParsedDataValidation> = Vec::new();

    #[derive(Clone, Copy, Debug)]
    enum FormulaTarget {
        Formula1,
        Formula2,
    }

    struct CurrentValidation {
        ranges: Vec<Range>,
        validation: DataValidation,
    }

    let mut current: Option<CurrentValidation> = None;
    let mut in_formula: Option<FormulaTarget> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref() == b"dataValidation" => {
                // Defaults based on observed Excel behavior and the OOXML schema.
                let mut kind: Option<DataValidationKind> = None;
                let mut operator: Option<DataValidationOperator> = None;
                let mut allow_blank = false;
                let mut show_input_message = true;
                let mut show_error_message = true;
                // Model semantics: `show_drop_down=true` means show the in-cell dropdown arrow.
                let mut show_drop_down = true;
                let mut ranges: Vec<Range> = Vec::new();

                let mut prompt_title: Option<String> = None;
                let mut prompt: Option<String> = None;

                let mut error_style: Option<DataValidationErrorStyle> = None;
                let mut error_title: Option<String> = None;
                let mut error: Option<String> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"type" => kind = parse_kind(&val),
                        b"operator" => operator = parse_operator(&val),
                        b"allowBlank" => allow_blank = parse_xml_bool(&val),
                        b"showInputMessage" => show_input_message = parse_xml_bool(&val),
                        b"showErrorMessage" => show_error_message = parse_xml_bool(&val),
                        b"showDropDown" => {
                            // OOXML `showDropDown` is inverted ("hide the dropdown").
                            show_drop_down = !parse_xml_bool(&val)
                        }
                        b"sqref" => {
                            ranges =
                                parse_sqref(&val).map_err(|e| XlsxError::Invalid(e.to_string()))?
                        }
                        b"promptTitle" => prompt_title = Some(val),
                        b"prompt" => prompt = Some(val),
                        b"errorStyle" => error_style = parse_error_style(&val),
                        b"errorTitle" => error_title = Some(val),
                        b"error" => error = Some(val),
                        _ => {}
                    }
                }

                let Some(kind) = kind else {
                    // Skip unsupported/disabled validations (`type="none"` or unknown).
                    buf.clear();
                    continue;
                };

                let input_message = if prompt_title.is_some() || prompt.is_some() {
                    Some(DataValidationInputMessage {
                        title: prompt_title,
                        body: prompt,
                    })
                } else {
                    None
                };

                let error_alert =
                    if error_style.is_some() || error_title.is_some() || error.is_some() {
                        Some(DataValidationErrorAlert {
                            style: error_style.unwrap_or_default(),
                            title: error_title,
                            body: error,
                        })
                    } else {
                        None
                    };

                current = Some(CurrentValidation {
                    ranges,
                    validation: DataValidation {
                        kind,
                        operator,
                        formula1: String::new(),
                        formula2: None,
                        allow_blank,
                        show_input_message,
                        show_error_message,
                        show_drop_down,
                        input_message,
                        error_alert,
                    },
                });
            }
            Event::Empty(e) if e.local_name().as_ref() == b"dataValidation" => {
                // Self-closing `<dataValidation/>` elements can occur (though uncommon).
                let mut kind: Option<DataValidationKind> = None;
                let mut operator: Option<DataValidationOperator> = None;
                let mut allow_blank = false;
                let mut show_input_message = true;
                let mut show_error_message = true;
                let mut show_drop_down = true;
                let mut ranges: Vec<Range> = Vec::new();

                let mut prompt_title: Option<String> = None;
                let mut prompt: Option<String> = None;

                let mut error_style: Option<DataValidationErrorStyle> = None;
                let mut error_title: Option<String> = None;
                let mut error: Option<String> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"type" => kind = parse_kind(&val),
                        b"operator" => operator = parse_operator(&val),
                        b"allowBlank" => allow_blank = parse_xml_bool(&val),
                        b"showInputMessage" => show_input_message = parse_xml_bool(&val),
                        b"showErrorMessage" => show_error_message = parse_xml_bool(&val),
                        b"showDropDown" => show_drop_down = !parse_xml_bool(&val),
                        b"sqref" => {
                            ranges =
                                parse_sqref(&val).map_err(|e| XlsxError::Invalid(e.to_string()))?
                        }
                        b"promptTitle" => prompt_title = Some(val),
                        b"prompt" => prompt = Some(val),
                        b"errorStyle" => error_style = parse_error_style(&val),
                        b"errorTitle" => error_title = Some(val),
                        b"error" => error = Some(val),
                        _ => {}
                    }
                }

                let Some(kind) = kind else {
                    buf.clear();
                    continue;
                };

                let input_message = if prompt_title.is_some() || prompt.is_some() {
                    Some(DataValidationInputMessage {
                        title: prompt_title,
                        body: prompt,
                    })
                } else {
                    None
                };

                let error_alert =
                    if error_style.is_some() || error_title.is_some() || error.is_some() {
                        Some(DataValidationErrorAlert {
                            style: error_style.unwrap_or_default(),
                            title: error_title,
                            body: error,
                        })
                    } else {
                        None
                    };

                out.push(ParsedDataValidation {
                    ranges,
                    validation: DataValidation {
                        kind,
                        operator,
                        formula1: String::new(),
                        formula2: None,
                        allow_blank,
                        show_input_message,
                        show_error_message,
                        show_drop_down,
                        input_message,
                        error_alert,
                    },
                });
            }
            Event::Start(e) if e.local_name().as_ref() == b"formula1" => {
                if current.is_some() {
                    in_formula = Some(FormulaTarget::Formula1);
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"formula1" => {
                drop(e);
                // Leave as empty string.
            }
            Event::Start(e) if e.local_name().as_ref() == b"formula2" => {
                if current.is_some() {
                    in_formula = Some(FormulaTarget::Formula2);
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"formula2" => {
                if let Some(cur) = current.as_mut() {
                    cur.validation.formula2 = Some(String::new());
                }
            }
            Event::Text(e) => {
                let Some(target) = in_formula else {
                    buf.clear();
                    continue;
                };
                let Some(cur) = current.as_mut() else {
                    buf.clear();
                    continue;
                };
                let text = e.unescape()?.into_owned();
                let normalized = strip_leading_equals(&text).to_string();
                match target {
                    FormulaTarget::Formula1 => cur.validation.formula1.push_str(&normalized),
                    FormulaTarget::Formula2 => {
                        cur.validation
                            .formula2
                            .get_or_insert_with(String::new)
                            .push_str(&normalized);
                    }
                }
            }
            Event::CData(e) => {
                let Some(target) = in_formula else {
                    buf.clear();
                    continue;
                };
                let Some(cur) = current.as_mut() else {
                    buf.clear();
                    continue;
                };
                let text = String::from_utf8_lossy(e.as_ref()).into_owned();
                let normalized = strip_leading_equals(&text).to_string();
                match target {
                    FormulaTarget::Formula1 => cur.validation.formula1.push_str(&normalized),
                    FormulaTarget::Formula2 => {
                        cur.validation
                            .formula2
                            .get_or_insert_with(String::new)
                            .push_str(&normalized);
                    }
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"formula1" => {
                drop(e);
                in_formula = None;
            }
            Event::End(e) if e.local_name().as_ref() == b"formula2" => {
                drop(e);
                in_formula = None;
            }
            Event::End(e) if e.local_name().as_ref() == b"dataValidation" => {
                drop(e);
                if let Some(mut cur) = current.take() {
                    // Excel stores formulas without a leading '='; normalize.
                    cur.validation.formula1 =
                        strip_leading_equals(&cur.validation.formula1).to_string();
                    if let Some(f2) = cur.validation.formula2.as_deref() {
                        let normalized = strip_leading_equals(f2).to_string();
                        cur.validation.formula2 = Some(normalized);
                    }
                    out.push(ParsedDataValidation {
                        ranges: cur.ranges,
                        validation: cur.validation,
                    });
                }
                in_formula = None;
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}
