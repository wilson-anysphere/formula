use crate::merge_cells::write_merge_cells_section;
use formula_model::{Alignment, Range};
use std::io::{Cursor, Write};
use thiserror::Error;
use zip::write::FileOptions;

#[derive(Debug, Error)]
pub enum MinimalXlsxError {
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
}

/// Write a minimal XLSX workbook that contains a single worksheet with merged cells.
///
/// This is a targeted serializer used for merge-cell round-trip tests.
pub fn write_minimal_xlsx(
    merges: &[Range],
    alignments: &[Alignment],
) -> Result<Vec<u8>, MinimalXlsxError> {
    let mut buffer = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut buffer);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("[Content_Types].xml", options)?;
        zip.write_all(content_types_xml().as_bytes())?;

        zip.start_file("_rels/.rels", options)?;
        zip.write_all(rels_xml().as_bytes())?;

        zip.start_file("xl/workbook.xml", options)?;
        zip.write_all(workbook_xml().as_bytes())?;

        zip.start_file("xl/_rels/workbook.xml.rels", options)?;
        zip.write_all(workbook_rels_xml().as_bytes())?;

        zip.start_file("xl/styles.xml", options)?;
        zip.write_all(styles_xml(alignments).as_bytes())?;

        zip.start_file("xl/worksheets/sheet1.xml", options)?;
        zip.write_all(worksheet_xml(merges).as_bytes())?;

        zip.finish()?;
    }
    Ok(buffer.into_inner())
}

fn content_types_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"#
    .to_owned()
}

fn rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#
    .to_owned()
}

fn workbook_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#
    .to_owned()
}

fn workbook_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"#
    .to_owned()
}

fn styles_xml(alignments: &[Alignment]) -> String {
    let mut out = String::new();
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push('\n');
    out.push_str(r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#);
    out.push('\n');
    out.push_str(
        r#"  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
    );
    out.push('\n');
    out.push_str(r#"  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>"#);
    out.push('\n');
    out.push_str(r#"  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#);
    out.push('\n');
    out.push_str(r#"  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#);
    out.push('\n');
    out.push_str(&format!(r#"  <cellXfs count="{}">"#, alignments.len().max(1)));
    out.push('\n');

    // Always include index 0 default.
    let alignments = if alignments.is_empty() {
        &[Alignment::default()][..]
    } else {
        alignments
    };

    for (_idx, alignment) in alignments.iter().enumerate() {
        let is_default = alignment == &Alignment::default();
        if is_default {
            out.push_str(r#"    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#);
            out.push('\n');
            continue;
        }

        out.push_str(r#"    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">"#);
        out.push('\n');
        out.push_str("      ");
        out.push_str(&alignment_xml(alignment));
        out.push('\n');
        out.push_str("    </xf>\n");
    }

    out.push_str("  </cellXfs>\n");
    out.push_str(r#"  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#);
    out.push('\n');
    out.push_str("</styleSheet>\n");
    out
}

fn alignment_xml(alignment: &Alignment) -> String {
    let mut out = String::new();
    out.push_str("<alignment");
    if let Some(h) = alignment.horizontal {
        let h = match h {
            formula_model::HorizontalAlignment::General => "general",
            formula_model::HorizontalAlignment::Left => "left",
            formula_model::HorizontalAlignment::Center => "center",
            formula_model::HorizontalAlignment::Right => "right",
        };
        out.push_str(&format!(r#" horizontal="{}""#, h));
    }
    if let Some(v) = alignment.vertical {
        let v = match v {
            formula_model::VerticalAlignment::Top => "top",
            formula_model::VerticalAlignment::Center => "center",
            formula_model::VerticalAlignment::Bottom => "bottom",
        };
        out.push_str(&format!(r#" vertical="{}""#, v));
    }
    if alignment.wrap_text {
        out.push_str(r#" wrapText="1""#);
    }
    if alignment.text_rotation != 0 {
        out.push_str(&format!(r#" textRotation="{}""#, alignment.text_rotation));
    }
    out.push_str("/>");
    out
}

fn worksheet_xml(merges: &[Range]) -> String {
    let mut out = String::new();
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push('\n');
    out.push_str(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#);
    out.push('\n');

    out.push_str("  <sheetData>\n");
    out.push_str(r#"    <row r="1">"#);
    out.push('\n');
    out.push_str(r#"      <c r="A1" t="inlineStr" s="1"><is><t>Merged</t></is></c>"#);
    out.push('\n');
    out.push_str("    </row>\n");
    out.push_str("  </sheetData>\n");

    if !merges.is_empty() {
        out.push_str("  ");
        out.push_str(&write_merge_cells_section(merges));
    }

    out.push_str("</worksheet>\n");
    out
}
