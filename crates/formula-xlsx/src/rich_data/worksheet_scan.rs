use std::io::Cursor;

use formula_model::CellRef;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::local_name;
use crate::XlsxError;

use super::RichDataError;

/// Scan a worksheet XML part for cells that reference `valueMetadata` (`vm`) or `cellMetadata`
/// (`cm`) records.
///
/// This uses `quick_xml` streaming parsing to avoid materializing a full DOM for very large sheets.
pub fn scan_cells_with_metadata_indices(
    sheet_xml: &[u8],
) -> Result<Vec<(CellRef, Option<u32>, Option<u32>)>, RichDataError> {
    let mut reader = Reader::from_reader(Cursor::new(sheet_xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_sheet_data = false;
    let mut out: Vec<(CellRef, Option<u32>, Option<u32>)> = Vec::new();

    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(XlsxError::from)?
        {
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = true;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                // Empty `<sheetData/>` - nothing to scan.
                in_sheet_data = false;
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = false;
            }

            Event::Start(ref e) | Event::Empty(ref e)
                if in_sheet_data && local_name(e.name().as_ref()) == b"c" =>
            {
                let mut r: Option<std::borrow::Cow<'_, str>> = None;
                let mut vm: Option<u32> = None;
                let mut cm: Option<u32> = None;

                for attr in e.attributes() {
                    let attr = attr.map_err(XlsxError::from)?;
                    match attr.key.as_ref() {
                        b"r" => r = Some(attr.unescape_value().map_err(XlsxError::from)?),
                        b"vm" => {
                            let v = attr.unescape_value().map_err(XlsxError::from)?;
                            vm = v.as_ref().trim().parse::<u32>().ok();
                        }
                        b"cm" => {
                            let v = attr.unescape_value().map_err(XlsxError::from)?;
                            cm = v.as_ref().trim().parse::<u32>().ok();
                        }
                        _ => {}
                    }
                }

                if vm.is_some() || cm.is_some() {
                    let Some(r) = r else {
                        continue;
                    };
                    let Ok(cell_ref) = CellRef::from_a1(r.as_ref()) else {
                        continue;
                    };
                    out.push((cell_ref, vm, cm));
                }
            }

            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scan_cells_with_metadata_indices_skips_plain_cells_in_large_sheet() {
        // Use a prefixed SpreadsheetML namespace to ensure we match on local name only.
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        );
        xml.push_str("<x:sheetData>");

        // 10k plain cells (1000 rows x 10 cols) + a few metadata-indexed ones.
        for row_1 in 1..=1000u32 {
            xml.push_str(&format!(r#"<x:row r="{row_1}">"#));

            for (col_letter, col_idx) in ('A'..='J').zip(0u32..) {
                let a1 = format!("{col_letter}{row_1}");
                match a1.as_str() {
                    "B2" => xml.push_str(r#"<x:c r="B2" vm="1"/>"#),
                    "D500" => xml.push_str(r#"<x:c r="D500" cm="7"/>"#),
                    "J1000" => xml.push_str(r#"<x:c r="J1000" vm="42" cm="3"/>"#),
                    _ => {
                        // Mix empty and non-empty cell representations to exercise both
                        // `Event::Empty` and `Event::Start`.
                        if col_idx % 2 == 0 {
                            xml.push_str(&format!(r#"<x:c r="{a1}"/>"#));
                        } else {
                            xml.push_str(&format!(r#"<x:c r="{a1}"><x:v>0</x:v></x:c>"#));
                        }
                    }
                }
            }

            xml.push_str("</x:row>");
        }

        xml.push_str("</x:sheetData></x:worksheet>");

        let found = scan_cells_with_metadata_indices(xml.as_bytes()).expect("scan worksheet");
        assert_eq!(found.len(), 3);

        assert_eq!(
            found[0],
            (CellRef::from_a1("B2").unwrap(), Some(1), None)
        );
        assert_eq!(
            found[1],
            (CellRef::from_a1("D500").unwrap(), None, Some(7))
        );
        assert_eq!(
            found[2],
            (
                CellRef::from_a1("J1000").unwrap(),
                Some(42),
                Some(3)
            )
        );
    }
}
