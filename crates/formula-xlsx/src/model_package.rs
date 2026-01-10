use std::fs;
use std::path::Path;

use formula_model::{Cell, CellRef, CellValue, Workbook, Worksheet, WorksheetId};

use crate::package::{XlsxError, XlsxPackage};
use crate::path::resolve_target;
use crate::styles::{StylesPart, StylesPartError};
use crate::xml::{XmlDomError, XmlElement, XmlNode};

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_TYPE_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

const DEFAULT_STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"#;

#[derive(Debug, thiserror::Error)]
pub enum WorkbookPackageError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Package(#[from] XlsxError),
    #[error("missing part: {0}")]
    MissingPart(String),
    #[error("xml error: {part}: {source}")]
    Xml {
        part: String,
        #[source]
        source: XmlDomError,
    },
    #[error(transparent)]
    Styles(#[from] StylesPartError),
    #[error("invalid cell reference {reference}: {source}")]
    InvalidCellReference {
        reference: String,
        #[source]
        source: formula_model::A1ParseError,
    },
    #[error("workbook.xml missing <sheets> element")]
    MissingSheets,
    #[error("sheet element missing required attribute: {0}")]
    MissingSheetAttribute(&'static str),
    #[error("missing worksheet relationship target for sheet {0}")]
    MissingSheetRelationship(String),
}

#[derive(Debug, Clone)]
struct Relationship {
    id: String,
    type_: String,
    target: String,
}

#[derive(Debug, Clone)]
struct WorksheetPart {
    part_name: String,
    sheet_id: WorksheetId,
    xml: XmlElement,
}

/// A workbook model layered over an [`XlsxPackage`].
///
/// This is intentionally minimal: today it focuses on the cell style pipeline (`styles.xml`
/// + worksheet `s` attributes) and preserves all other parts verbatim.
#[derive(Debug)]
pub struct WorkbookPackage {
    pub workbook: Workbook,

    package: XlsxPackage,
    styles_part_name: String,
    styles: StylesPart,
    worksheets: Vec<WorksheetPart>,
}

impl WorkbookPackage {
    pub fn load(path: &Path) -> Result<Self, WorkbookPackageError> {
        let bytes = fs::read(path)?;
        let package = XlsxPackage::from_bytes(&bytes)?;

        let workbook_part = "xl/workbook.xml";
        let workbook_xml = package
            .part(workbook_part)
            .ok_or_else(|| WorkbookPackageError::MissingPart(workbook_part.to_string()))?;
        let workbook_root = XmlElement::parse(workbook_xml).map_err(|source| WorkbookPackageError::Xml {
            part: workbook_part.to_string(),
            source,
        })?;

        let rels_part = "xl/_rels/workbook.xml.rels";
        let rels_xml = package
            .part(rels_part)
            .ok_or_else(|| WorkbookPackageError::MissingPart(rels_part.to_string()))?;
        let rels_root = XmlElement::parse(rels_xml).map_err(|source| WorkbookPackageError::Xml {
            part: rels_part.to_string(),
            source,
        })?;

        let rels = parse_relationships(&rels_root);
        let styles_part_name = rels
            .iter()
            .find(|rel| rel.type_ == REL_TYPE_STYLES)
            .map(|rel| resolve_target(workbook_part, &rel.target))
            .unwrap_or_else(|| "xl/styles.xml".to_string());

        let mut workbook = Workbook::new();
        let styles = match package.part(&styles_part_name) {
            Some(bytes) => StylesPart::parse(bytes, &mut workbook.styles)?,
            None => StylesPart::parse(DEFAULT_STYLES_XML.as_bytes(), &mut workbook.styles)?,
        };

        let Some(sheets_el) = workbook_root.child("sheets") else {
            return Err(WorkbookPackageError::MissingSheets);
        };

        let mut worksheets = Vec::new();
        for sheet_el in sheets_el.children.iter().filter_map(|n| match n {
            XmlNode::Element(el) if el.name.local == "sheet" => Some(el),
            _ => None,
        }) {
            let name = sheet_el
                .attr("name")
                .ok_or(WorkbookPackageError::MissingSheetAttribute("name"))?
                .to_string();
            let rel_id = sheet_el
                .attr_ns(REL_NS, "id")
                .ok_or(WorkbookPackageError::MissingSheetAttribute("r:id"))?;

            let target = rels
                .iter()
                .find(|rel| rel.id == rel_id && rel.type_ == REL_TYPE_WORKSHEET)
                .map(|rel| rel.target.clone())
                .ok_or_else(|| WorkbookPackageError::MissingSheetRelationship(name.clone()))?;

            let sheet_part_name = resolve_target(workbook_part, &target);
            let sheet_xml = package
                .part(&sheet_part_name)
                .ok_or_else(|| WorkbookPackageError::MissingPart(sheet_part_name.clone()))?;
            let sheet_root =
                XmlElement::parse(sheet_xml).map_err(|source| WorkbookPackageError::Xml {
                    part: sheet_part_name.clone(),
                    source,
                })?;

            let sheet_id = workbook.add_sheet(name.clone());
            let sheet_model = workbook.sheet_mut(sheet_id).expect("sheet just added");
            parse_sheet_cells(sheet_model, &sheet_root, &styles)?;

            worksheets.push(WorksheetPart {
                part_name: sheet_part_name,
                sheet_id,
                xml: sheet_root,
            });
        }

        Ok(Self {
            workbook,
            package,
            styles_part_name,
            styles,
            worksheets,
        })
    }

    pub fn save(&mut self, out_path: &Path) -> Result<(), WorkbookPackageError> {
        // Ensure we have xf indices for any styles referenced by stored cells.
        for sheet in &self.workbook.sheets {
            for (_, cell) in sheet.iter_cells() {
                self.styles
                    .xf_index_for_style(cell.style_id, &self.workbook.styles)?;
            }
        }

        // Update worksheet `s` attributes from style_ids.
        for part in &mut self.worksheets {
            let Some(sheet) = self.workbook.sheet(part.sheet_id) else {
                continue;
            };
            update_sheet_styles(part, sheet, &mut self.styles, &self.workbook.styles)?;
        }

        // Replace styles.xml.
        self.package
            .set_part(self.styles_part_name.clone(), self.styles.to_xml_bytes());

        // Replace worksheet parts.
        for part in &self.worksheets {
            self.package.set_part(
                part.part_name.clone(),
                part.xml.to_xml_string().into_bytes(),
            );
        }

        let bytes = self.package.write_to_bytes()?;
        fs::write(out_path, bytes)?;
        Ok(())
    }

    pub fn styles(&self) -> &StylesPart {
        &self.styles
    }

    pub fn styles_mut(&mut self) -> &mut StylesPart {
        &mut self.styles
    }

    pub fn xf_index_for_style(&mut self, style_id: u32) -> Result<u32, WorkbookPackageError> {
        Ok(self
            .styles
            .xf_index_for_style(style_id, &self.workbook.styles)?)
    }
}

fn parse_relationships(root: &XmlElement) -> Vec<Relationship> {
    let mut rels = Vec::new();
    for rel in root.children.iter().filter_map(|n| match n {
        XmlNode::Element(el) if el.name.local == "Relationship" => Some(el),
        _ => None,
    }) {
        let id = rel.attr("Id").unwrap_or_default().to_string();
        let type_ = rel.attr("Type").unwrap_or_default().to_string();
        let target = rel.attr("Target").unwrap_or_default().to_string();
        if !id.is_empty() && !type_.is_empty() && !target.is_empty() {
            rels.push(Relationship { id, type_, target });
        }
    }
    rels
}

fn parse_sheet_cells(
    sheet: &mut Worksheet,
    root: &XmlElement,
    styles: &StylesPart,
) -> Result<(), WorkbookPackageError> {
    // We only need a subset of sheet parsing for fixtures today: cells under sheetData/row/c.
    let Some(sheet_data) = root.child("sheetData") else {
        return Ok(());
    };

    for row in sheet_data.children.iter().filter_map(|n| match n {
        XmlNode::Element(el) if el.name.local == "row" => Some(el),
        _ => None,
    }) {
        for c in row.children.iter().filter_map(|n| match n {
            XmlNode::Element(el) if el.name.local == "c" => Some(el),
            _ => None,
        }) {
            let Some(r) = c.attr("r") else {
                continue;
            };
            let cell_ref = CellRef::from_a1(r).map_err(|source| WorkbookPackageError::InvalidCellReference {
                reference: r.to_string(),
                source,
            })?;

            let style_id = c
                .attr("s")
                .and_then(|v| v.parse::<u32>().ok())
                .map(|xf| styles.style_id_for_xf(xf))
                .unwrap_or(0);

            let cell_type = c.attr("t").unwrap_or_default();

            let mut cell = Cell::default();
            cell.style_id = style_id;

            if let Some(f) = c.child("f").and_then(|f| f.text()) {
                cell.formula = Some(f.to_string());
            }

            match cell_type {
                "inlineStr" => {
                    if let Some(text) = c
                        .child("is")
                        .and_then(|is| is.child("t"))
                        .and_then(|t| t.text())
                    {
                        cell.value = CellValue::String(text.to_string());
                    }
                }
                "b" => {
                    if let Some(v) = c.child("v").and_then(|v| v.text()) {
                        cell.value = CellValue::Boolean(v == "1");
                    }
                }
                _ => {
                    if let Some(v) = c.child("v").and_then(|v| v.text()) {
                        if let Ok(num) = v.parse::<f64>() {
                            cell.value = CellValue::Number(num);
                        } else {
                            cell.value = CellValue::String(v.to_string());
                        }
                    }
                }
            }

            sheet.set_cell(cell_ref, cell);
        }
    }

    Ok(())
}

fn update_sheet_styles(
    part: &mut WorksheetPart,
    sheet: &Worksheet,
    styles: &mut StylesPart,
    style_table: &formula_model::StyleTable,
) -> Result<(), WorkbookPackageError> {
    let Some(sheet_data) = part.xml.child_mut("sheetData") else {
        return Ok(());
    };

    for row in sheet_data.children.iter_mut().filter_map(|n| match n {
        XmlNode::Element(el) if el.name.local == "row" => Some(el),
        _ => None,
    }) {
        for c in row.children.iter_mut().filter_map(|n| match n {
            XmlNode::Element(el) if el.name.local == "c" => Some(el),
            _ => None,
        }) {
            let Some(r) = c.attr("r").map(|s| s.to_string()) else {
                continue;
            };
            let cell_ref = match CellRef::from_a1(&r) {
                Ok(r) => r,
                Err(_) => continue,
            };

            let style_id = sheet.cell(cell_ref).map(|c| c.style_id).unwrap_or(0);
            let xf_index = styles.xf_index_for_style(style_id, style_table)?;

            if xf_index == 0 {
                c.remove_attr("s");
            } else {
                c.set_attr("s", xf_index.to_string());
            }
        }
    }
    Ok(())
}
