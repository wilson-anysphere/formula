use formula_model::Range;
use std::fmt::Write as _;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::XlsxError;

fn insert_before_tag(name: &[u8]) -> bool {
    matches!(
        name,
        // Elements that come after <mergeCells> in the SpreadsheetML schema.
        b"phoneticPr"
            | b"conditionalFormatting"
            | b"dataValidations"
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
    )
}

#[must_use]
pub fn write_merge_cells_section(merges: &[Range]) -> String {
    if merges.is_empty() {
        return String::new();
    }

    let mut out = String::new();
    let _ = writeln!(out, r#"<mergeCells count="{}">"#, merges.len());
    for merge in merges {
        let _ = writeln!(
            out,
            r#"  <mergeCell ref="{}"/>"#,
            merge
        );
    }
    out.push_str("</mergeCells>\n");
    out
}

/// Write a minimal worksheet XML containing sheet data and merge cells.
///
/// This is *not* a full-fidelity round-trip serializer; it is intended for tests and
/// for the merge-cells subsystem in isolation.
#[must_use]
pub fn write_worksheet_xml(merges: &[Range]) -> String {
    let mut out = String::new();
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push('\n');
    out.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    );
    out.push('\n');
    out.push_str("  <sheetData/>\n");
    if !merges.is_empty() {
        out.push_str("  ");
        out.push_str(&write_merge_cells_section(merges).replace('\n', "\n  "));
        // `replace` adds trailing indentation; normalize.
        out = out.replace("\n  </mergeCells>", "\n</mergeCells>");
        if !out.ends_with('\n') {
            out.push('\n');
        }
    }
    out.push_str("</worksheet>\n");
    out
}

/// Update (or remove) a worksheet `<mergeCells>` section to match `merges`.
///
/// If the worksheet already contains `<mergeCells>`, it is replaced. If it does not
/// and `merges` is non-empty, the block is inserted before the end of the worksheet
/// (preferably before elements that are required to come after it, e.g.
/// `<conditionalFormatting>`, `<hyperlinks>`, `<pageMargins>`, `<tableParts>`, `<extLst>`).
pub fn update_worksheet_xml(sheet_xml: &str, merges: &[Range]) -> Result<String, XlsxError> {
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
                Event::End(_) => skip_depth -= 1,
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"mergeCells" => {
                replaced = true;
                if !merges.is_empty() {
                    write_merge_cells_block(&mut writer, merges, worksheet_prefix.as_deref())?;
                }
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"mergeCells" => {
                replaced = true;
                if !merges.is_empty() {
                    write_merge_cells_block(&mut writer, merges, worksheet_prefix.as_deref())?;
                }
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if !replaced
                    && !merges.is_empty()
                    && insert_before_tag(e.local_name().as_ref()) =>
            {
                write_merge_cells_block(&mut writer, merges, worksheet_prefix.as_deref())?;
                replaced = true;
                writer.write_event(event.to_owned())?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if !replaced && !merges.is_empty() {
                    write_merge_cells_block(&mut writer, merges, worksheet_prefix.as_deref())?;
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

fn write_merge_cells_block<W: std::io::Write>(
    writer: &mut Writer<W>,
    merges: &[Range],
    prefix: Option<&str>,
) -> Result<(), XlsxError> {
    let merge_cells_tag = crate::xml::prefixed_tag(prefix, "mergeCells");
    let merge_cell_tag = crate::xml::prefixed_tag(prefix, "mergeCell");
    let count = merges.len().to_string();
    let mut start = BytesStart::new(merge_cells_tag.as_str());
    start.push_attribute(("count", count.as_str()));
    writer.write_event(Event::Start(start))?;

    for merge in merges {
        let range = merge.to_string();
        let mut elem = BytesStart::new(merge_cell_tag.as_str());
        elem.push_attribute(("ref", range.as_str()));
        writer.write_event(Event::Empty(elem))?;
    }

    writer.write_event(Event::End(BytesEnd::new(merge_cells_tag.as_str())))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn update_inserts_before_table_parts_when_missing() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#;
        let merges = vec![Range::from_a1("A1:B2").unwrap()];
        let updated = update_worksheet_xml(xml, &merges).unwrap();

        let merge_pos = updated.find("<mergeCells").expect("mergeCells inserted");
        let table_pos = updated.find("<tableParts").expect("tableParts exists");
        assert!(
            merge_pos < table_pos,
            "expected mergeCells before tableParts, got:\n{updated}"
        );
    }

    #[test]
    fn update_inserts_before_page_margins_when_missing() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#;
        let merges = vec![Range::from_a1("A1:B2").unwrap()];
        let updated = update_worksheet_xml(xml, &merges).unwrap();

        let merge_pos = updated.find("<mergeCells").expect("mergeCells inserted");
        let margins_pos = updated.find("<pageMargins").expect("pageMargins exists");
        assert!(
            merge_pos < margins_pos,
            "expected mergeCells before pageMargins, got:\n{updated}"
        );
    }
}
