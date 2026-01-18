use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use formula_model::{HiddenState, Outline, OutlineEntry, OutlinePr};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use thiserror::Error;

use crate::{XlsxError, XlsxPackage};

#[derive(Debug, Error)]
pub enum OutlineXlsxError {
    #[error(transparent)]
    Package(#[from] XlsxError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("utf8 error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("missing worksheet entry: {0}")]
    MissingWorksheet(String),
    #[error("invalid xml attribute value for {0}: {1}")]
    InvalidAttr(&'static str, String),
}

/// Reads outline metadata from a worksheet XML document (`xl/worksheets/sheetN.xml`).
pub fn read_outline_from_worksheet_xml(xml: &str) -> Result<Outline, OutlineXlsxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut outline = Outline::default();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => match e.local_name().as_ref() {
                b"outlinePr" => parse_outline_pr(&mut outline.pr, &e)?,
                b"row" => parse_row_outline(&mut outline, &e)?,
                b"col" => parse_col_outline(&mut outline, &e)?,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    outline.recompute_outline_hidden_rows();
    outline.recompute_outline_hidden_cols();

    // Heuristic: if the row/col is hidden and we believe it's hidden by outline,
    // clear the user flag so expand restores visibility.
    for (_, entry) in outline.rows.iter_mut() {
        if entry.hidden.outline {
            entry.hidden.user = false;
        }
    }
    for (_, entry) in outline.cols.iter_mut() {
        if entry.hidden.outline {
            entry.hidden.user = false;
        }
    }

    Ok(outline)
}

fn parse_outline_pr(pr: &mut OutlinePr, e: &BytesStart<'_>) -> Result<(), OutlineXlsxError> {
    for attr in e.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"summaryBelow" => pr.summary_below = parse_bool(attr.value.as_ref())?,
            b"summaryRight" => pr.summary_right = parse_bool(attr.value.as_ref())?,
            b"showOutlineSymbols" => pr.show_outline_symbols = parse_bool(attr.value.as_ref())?,
            _ => {}
        }
    }
    Ok(())
}

fn parse_row_outline(outline: &mut Outline, e: &BytesStart<'_>) -> Result<(), OutlineXlsxError> {
    let mut row_index: Option<u32> = None;
    let mut entry = OutlineEntry::default();
    for attr in e.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"r" => row_index = Some(parse_u32(attr.value.as_ref(), "r")?),
            b"outlineLevel" => entry.level = parse_u8(attr.value.as_ref(), "outlineLevel")?,
            b"hidden" => entry.hidden.user = parse_bool(attr.value.as_ref())?,
            b"collapsed" => entry.collapsed = parse_bool(attr.value.as_ref())?,
            _ => {}
        }
    }
    if let Some(index) = row_index {
        // Only store non-default entries so `Outline` stays compact (and so sheets without any
        // outline metadata keep `Outline::default()`).
        if entry != OutlineEntry::default() {
            let stored = outline.rows.entry_mut(index);
            stored.level = entry.level;
            stored.collapsed = entry.collapsed;
            stored.hidden.user = entry.hidden.user;
        }
    }
    Ok(())
}

fn parse_col_outline(outline: &mut Outline, e: &BytesStart<'_>) -> Result<(), OutlineXlsxError> {
    let mut min: Option<u32> = None;
    let mut max: Option<u32> = None;
    let mut entry = OutlineEntry::default();
    for attr in e.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"min" => min = Some(parse_u32(attr.value.as_ref(), "min")?),
            b"max" => max = Some(parse_u32(attr.value.as_ref(), "max")?),
            b"outlineLevel" => entry.level = parse_u8(attr.value.as_ref(), "outlineLevel")?,
            b"hidden" => entry.hidden.user = parse_bool(attr.value.as_ref())?,
            b"collapsed" => entry.collapsed = parse_bool(attr.value.as_ref())?,
            _ => {}
        }
    }
    let Some(min) = min else { return Ok(()); };
    let Some(max) = max else { return Ok(()); };
    // Only store non-default entries so `Outline` stays compact (and so sheets without any outline
    // metadata keep `Outline::default()`).
    if entry != OutlineEntry::default() {
        for index in min..=max {
            let stored = outline.cols.entry_mut(index);
            stored.level = entry.level;
            stored.collapsed = entry.collapsed;
            stored.hidden.user = entry.hidden.user;
        }
    }
    Ok(())
}

fn parse_bool(value: &[u8]) -> Result<bool, OutlineXlsxError> {
    match value {
        b"1" | b"true" => Ok(true),
        b"0" | b"false" => Ok(false),
        other => Err(OutlineXlsxError::InvalidAttr(
            "bool",
            String::from_utf8_lossy(other).to_string(),
        )),
    }
}

fn parse_u32(value: &[u8], name: &'static str) -> Result<u32, OutlineXlsxError> {
    std::str::from_utf8(value)
        .ok()
        .and_then(|s| s.parse::<u32>().ok())
        .ok_or_else(|| OutlineXlsxError::InvalidAttr(name, String::from_utf8_lossy(value).into()))
}

fn parse_u8(value: &[u8], name: &'static str) -> Result<u8, OutlineXlsxError> {
    std::str::from_utf8(value)
        .ok()
        .and_then(|s| s.parse::<u8>().ok())
        .ok_or_else(|| OutlineXlsxError::InvalidAttr(name, String::from_utf8_lossy(value).into()))
}

/// Writes outline metadata back into a worksheet XML document.
///
/// This function preserves the original XML structure as much as possible by
/// streaming events through `quick-xml` and updating only outline-related
/// attributes.
pub fn write_outline_to_worksheet_xml(original_xml: &str, outline: &Outline) -> Result<String, OutlineXlsxError> {
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(original_xml)?;
    let mut reader = Reader::from_str(original_xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().expand_empty_elements = true;

    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let mut buf = Vec::new();
    let mut in_sheet_pr = false;
    let mut wrote_outline_pr = false;
    let mut sheet_pr_prefix: Option<String> = None;
    let mut skipping_cols_depth: Option<usize> = None;
    let mut cols_written = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(e) => {
                let name = e.local_name();
                if let Some(depth) = skipping_cols_depth {
                    skipping_cols_depth = Some(depth.saturating_add(1));
                } else if name.as_ref() == b"cols" && !outline.cols.is_empty() {
                    // Replace the entire <cols> section.
                    let cols_name = e.name();
                    let cols_name = cols_name.as_ref();
                    let cols_prefix = cols_name
                        .iter()
                        .rposition(|b| *b == b':')
                        .map(|idx| &cols_name[..idx])
                        .and_then(|p| std::str::from_utf8(p).ok());
                    write_cols(&mut writer, outline, cols_prefix)?;
                    cols_written = true;
                    skipping_cols_depth = Some(1);
                } else if name.as_ref() == b"sheetPr" {
                    in_sheet_pr = true;
                    wrote_outline_pr = false;
                    let sheet_pr_name = e.name();
                    let sheet_pr_name = sheet_pr_name.as_ref();
                    sheet_pr_prefix = sheet_pr_name
                        .iter()
                        .rposition(|b| *b == b':')
                        .map(|idx| &sheet_pr_name[..idx])
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    writer.write_event(Event::Start(e))?;
                } else if name.as_ref() == b"outlinePr" {
                    wrote_outline_pr = true;
                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Start(build_outline_pr(tag.as_str(), outline)))?;
                } else if name.as_ref() == b"sheetData" {
                    if !outline.cols.is_empty() && !cols_written {
                        write_cols(&mut writer, outline, worksheet_prefix.as_deref())?;
                        cols_written = true;
                    }
                    writer.write_event(Event::Start(e))?;
                } else if name.as_ref() == b"row" {
                    writer.write_event(Event::Start(update_row_attrs(e, outline)?))?;
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Event::Empty(e) => {
                let name = e.local_name();
                if skipping_cols_depth.is_some() {
                    // inside skipped <cols>, ignore
                } else if name.as_ref() == b"cols" && !outline.cols.is_empty() {
                    let cols_name = e.name();
                    let cols_name = cols_name.as_ref();
                    let cols_prefix = cols_name
                        .iter()
                        .rposition(|b| *b == b':')
                        .map(|idx| &cols_name[..idx])
                        .and_then(|p| std::str::from_utf8(p).ok());
                    write_cols(&mut writer, outline, cols_prefix)?;
                    cols_written = true;
                } else if name.as_ref() == b"sheetPr" {
                    // Expand <sheetPr/> so we can inject outlinePr.
                    let sheet_pr_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let sheet_pr_prefix = sheet_pr_tag
                        .split_once(':')
                        .map(|(p, _)| p.to_string());
                    let outline_pr_tag =
                        crate::xml::prefixed_tag(sheet_pr_prefix.as_deref(), "outlinePr");

                    writer.write_event(Event::Start(e.into_owned()))?;
                    writer.write_event(Event::Empty(build_outline_pr(
                        outline_pr_tag.as_str(),
                        outline,
                    )))?;
                    writer.write_event(Event::End(BytesEnd::new(sheet_pr_tag.as_str())))?;
                } else if name.as_ref() == b"outlinePr" {
                    wrote_outline_pr = true;
                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Empty(build_outline_pr(tag.as_str(), outline)))?;
                } else if name.as_ref() == b"sheetData" {
                    if !outline.cols.is_empty() && !cols_written {
                        write_cols(&mut writer, outline, worksheet_prefix.as_deref())?;
                        cols_written = true;
                    }
                    writer.write_event(Event::Empty(e))?;
                } else if name.as_ref() == b"row" {
                    writer.write_event(Event::Empty(update_row_attrs(e, outline)?))?;
                } else {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Event::End(e) => {
                let name = e.local_name();
                if let Some(depth) = skipping_cols_depth {
                    let next = depth - 1;
                    if next == 0 {
                        skipping_cols_depth = None;
                    } else {
                        skipping_cols_depth = Some(next);
                    }
                    // swallow the end event
                } else if name.as_ref() == b"sheetPr" {
                    if in_sheet_pr && !wrote_outline_pr {
                        let outline_pr_tag =
                            crate::xml::prefixed_tag(sheet_pr_prefix.as_deref(), "outlinePr");
                        writer.write_event(Event::Empty(build_outline_pr(
                            outline_pr_tag.as_str(),
                            outline,
                        )))?;
                        wrote_outline_pr = true;
                    }
                    in_sheet_pr = false;
                    sheet_pr_prefix = None;
                    writer.write_event(Event::End(e))?;
                } else if name.as_ref() == b"worksheet" {
                    writer.write_event(Event::End(e))?;
                } else {
                    writer.write_event(Event::End(e))?;
                }
            }
            Event::Text(t) => {
                if skipping_cols_depth.is_none() {
                    writer.write_event(Event::Text(t))?;
                }
            }
            Event::Comment(c) => {
                if skipping_cols_depth.is_none() {
                    writer.write_event(Event::Comment(c))?;
                }
            }
            Event::CData(c) => {
                if skipping_cols_depth.is_none() {
                    writer.write_event(Event::CData(c))?;
                }
            }
            Event::Decl(d) => writer.write_event(Event::Decl(d))?,
            Event::PI(p) => writer.write_event(Event::PI(p))?,
            Event::DocType(d) => writer.write_event(Event::DocType(d))?,
        }
        buf.clear();
    }

    let cursor = writer.into_inner();
    Ok(String::from_utf8(cursor.into_inner())?)
}

fn update_row_attrs<'a>(
    mut e: BytesStart<'a>,
    outline: &Outline,
) -> Result<BytesStart<'a>, OutlineXlsxError> {
    let mut row_index: Option<u32> = None;
    let mut attrs: Vec<(Cow<'static, [u8]>, Cow<'static, [u8]>)> = Vec::new();

    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
            row_index = Some(parse_u32(attr.value.as_ref(), "r")?);
        }

        // We'll rebuild the attribute list, omitting outline attrs for now.
        match attr.key.as_ref() {
            b"outlineLevel" | b"hidden" | b"collapsed" => {}
            _ => attrs.push((Cow::Owned(attr.key.as_ref().to_vec()), Cow::Owned(attr.value.to_vec()))),
        }
    }

    let index = row_index.unwrap_or(0);
    let entry = outline.rows.entry(index);

    if entry.level > 0 {
        attrs.push((
            Cow::Borrowed(b"outlineLevel"),
            Cow::Owned(entry.level.to_string().into_bytes()),
        ));
    }
    if entry.hidden.is_hidden() {
        attrs.push((Cow::Borrowed(b"hidden"), Cow::Borrowed(b"1")));
    }
    if entry.collapsed {
        attrs.push((Cow::Borrowed(b"collapsed"), Cow::Borrowed(b"1")));
    }

    // Clear attrs and re-add.
    e.clear_attributes();
    for (k, v) in attrs {
        e.push_attribute((k.as_ref(), v.as_ref()));
    }

    Ok(e)
}

fn build_outline_pr(tag: &str, outline: &Outline) -> BytesStart<'static> {
    let mut e = BytesStart::new(tag).into_owned();
    e.push_attribute(("summaryBelow", if outline.pr.summary_below { "1" } else { "0" }));
    e.push_attribute(("summaryRight", if outline.pr.summary_right { "1" } else { "0" }));
    e.push_attribute((
        "showOutlineSymbols",
        if outline.pr.show_outline_symbols { "1" } else { "0" },
    ));
    e
}

fn write_cols<W: Write>(
    writer: &mut Writer<W>,
    outline: &Outline,
    prefix: Option<&str>,
) -> Result<(), OutlineXlsxError> {
    if outline.cols.is_empty() {
        return Ok(());
    }

    let cols_tag = crate::xml::prefixed_tag(prefix, "cols");
    let col_tag = crate::xml::prefixed_tag(prefix, "col");

    writer.write_event(Event::Start(BytesStart::new(cols_tag.as_str())))?;

    // Group columns into contiguous ranges with identical outline attrs.
    let mut current_range: Option<(u32, u32, OutlineEntry)> = None;

    let mut sorted: BTreeMap<u32, OutlineEntry> = BTreeMap::new();
    for (index, entry) in outline.cols.iter() {
        sorted.insert(index, *entry);
    }

    for (index, entry) in sorted {
        let outline_entry = OutlineEntry {
            level: entry.level,
            collapsed: entry.collapsed,
            hidden: HiddenState {
                user: entry.hidden.user,
                outline: entry.hidden.outline,
                filter: entry.hidden.filter,
            },
        };

        match current_range {
            None => current_range = Some((index, index, outline_entry)),
            Some((start, end, current)) if end + 1 == index && current == outline_entry => {
                current_range = Some((start, index, current));
            }
            Some((start, end, current)) => {
                write_col(writer, col_tag.as_str(), start, end, &current)?;
                current_range = Some((index, index, outline_entry));
            }
        }
    }
    if let Some((start, end, current)) = current_range {
        write_col(writer, col_tag.as_str(), start, end, &current)?;
    }

    writer.write_event(Event::End(BytesEnd::new(cols_tag.as_str())))?;
    Ok(())
}

fn write_col<W: Write>(
    writer: &mut Writer<W>,
    tag: &str,
    min: u32,
    max: u32,
    entry: &OutlineEntry,
) -> Result<(), OutlineXlsxError> {
    let mut e = BytesStart::new(tag);
    let min_str = min.to_string();
    let max_str = max.to_string();
    let level_str = entry.level.to_string();
    e.push_attribute(("min", min_str.as_str()));
    e.push_attribute(("max", max_str.as_str()));

    if entry.level > 0 {
        e.push_attribute(("outlineLevel", level_str.as_str()));
    }
    if entry.hidden.is_hidden() {
        e.push_attribute(("hidden", "1"));
    }
    if entry.collapsed {
        e.push_attribute(("collapsed", "1"));
    }
    writer.write_event(Event::Empty(e))?;
    Ok(())
}

/// Reads a worksheet XML file from an XLSX package.
pub fn read_outline_from_xlsx_bytes(
    bytes: &[u8],
    worksheet_path: &str,
) -> Result<Outline, OutlineXlsxError> {
    let pkg = XlsxPackage::from_bytes(bytes)?;
    let Some(part) = pkg.part(worksheet_path) else {
        return Err(OutlineXlsxError::MissingWorksheet(
            worksheet_path.to_string(),
        ));
    };
    let xml = String::from_utf8(part.to_vec()).map_err(XlsxError::from)?;
    read_outline_from_worksheet_xml(&xml)
}

/// Writes outline metadata back into an XLSX package, replacing the worksheet XML at `worksheet_path`.
pub fn write_outline_to_xlsx_bytes(
    bytes: &[u8],
    worksheet_path: &str,
    outline: &Outline,
) -> Result<Vec<u8>, OutlineXlsxError> {
    let mut pkg = XlsxPackage::from_bytes(bytes)?;
    let Some(part) = pkg.part(worksheet_path) else {
        return Err(OutlineXlsxError::MissingWorksheet(
            worksheet_path.to_string(),
        ));
    };

    let original_xml = String::from_utf8(part.to_vec()).map_err(XlsxError::from)?;
    let updated_xml = write_outline_to_worksheet_xml(&original_xml, outline)?;
    pkg.set_part(worksheet_path.to_string(), updated_xml.into_bytes());
    Ok(pkg.write_to_bytes()?)
}
