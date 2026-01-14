#[cfg(not(target_arch = "wasm32"))]
use std::fs;
#[cfg(not(target_arch = "wasm32"))]
use std::path::Path;

use formula_model::{Cell, CellRef, CellValue, SheetVisibility, Workbook, Worksheet, WorksheetId};

use crate::package::{XlsxError, XlsxPackage};
use crate::path::{resolve_target, resolve_target_candidates};
use crate::styles::{StylesPart, StylesPartError};
use crate::zip_util::zip_part_names_equivalent;
use crate::xml::{XmlDomError, XmlElement, XmlNode};

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_TYPE_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_TYPE_CHARTSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

fn resolve_existing_part_name(
    package: &XlsxPackage,
    base_part: &str,
    target: &str,
) -> Option<String> {
    let candidates = resolve_target_candidates(base_part, target);
    // Prefer an exact match to keep part-name strings canonical when possible (some producers
    // percent-encode relationship targets while storing ZIP entry names unescaped, and vice versa).
    for candidate in &candidates {
        if candidate.is_empty() {
            continue;
        }
        if package.parts_map().contains_key(candidate) {
            return Some(candidate.clone());
        }
        let with_slash = format!("/{candidate}");
        if package.parts_map().contains_key(with_slash.as_str()) {
            return Some(with_slash);
        }
    }

    // Fall back to a linear scan for non-canonical producer output (case/leading slash/percent
    // encoding differences). This is only used in the `WorkbookPackage` pipeline, which loads a
    // bounded set of parts once, so the extra work is acceptable.
    for candidate in &candidates {
        if candidate.is_empty() {
            continue;
        }
        if let Some((name, _)) = package
            .parts_map()
            .iter()
            .find(|(key, _)| zip_part_names_equivalent(key.as_str(), candidate.as_str()))
        {
            return Some(name.clone());
        }
    }

    None
}

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
    #[error("invalid worksheet name: {0}")]
    InvalidSheetName(#[from] formula_model::SheetNameError),
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
    /// Load a workbook package from in-memory `.xlsx`/`.xlsm` bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, WorkbookPackageError> {
        let package = XlsxPackage::from_bytes(bytes)?;
        Self::from_package(package)
    }

    fn from_package(package: XlsxPackage) -> Result<Self, WorkbookPackageError> {
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
            .and_then(|rel| {
                if let Some(found) = resolve_existing_part_name(&package, workbook_part, &rel.target)
                {
                    return Some(found);
                }
                // Some broken workbooks omit or mis-point the styles relationship but still store
                // the canonical `xl/styles.xml` part. Fall back to it when present; otherwise, keep
                // the resolved target so round-tripping preserves the relationship path.
                resolve_existing_part_name(&package, "", "xl/styles.xml").or_else(|| {
                    resolve_target_candidates(workbook_part, &rel.target)
                        .into_iter()
                        .next()
                })
            })
            .or_else(|| resolve_existing_part_name(&package, "", "xl/styles.xml"))
            .unwrap_or_else(|| "xl/styles.xml".to_string());

        let mut workbook = Workbook::new();
        let styles =
            StylesPart::parse_or_default(package.part(&styles_part_name), &mut workbook.styles)?;

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
            let xlsx_sheet_id = sheet_el
                .attr("sheetId")
                .and_then(|v| v.parse::<u32>().ok());
            let visibility = match sheet_el.attr("state") {
                Some("hidden") => SheetVisibility::Hidden,
                Some("veryHidden") => SheetVisibility::VeryHidden,
                _ => SheetVisibility::Visible,
            };
            let rel_id = sheet_el
                .attr_ns(REL_NS, "id")
                .ok_or(WorkbookPackageError::MissingSheetAttribute("r:id"))?;

            let rel = rels
                .iter()
                .find(|rel| rel.id == rel_id)
                .ok_or_else(|| WorkbookPackageError::MissingSheetRelationship(name.clone()))?;

            // Preserve non-worksheet sheet types (chartsheets, etc) verbatim in the underlying
            // package without attempting to parse them into the workbook model.
            if rel.type_ == REL_TYPE_CHARTSHEET {
                continue;
            }

            if rel.type_ != REL_TYPE_WORKSHEET {
                continue;
            }

            let target = rel.target.clone();

            let sheet_part_name = resolve_existing_part_name(&package, workbook_part, &target)
                .unwrap_or_else(|| {
                    resolve_target_candidates(workbook_part, &target)
                        .into_iter()
                        .next()
                        .unwrap_or_else(|| resolve_target(workbook_part, &target))
                });
            let sheet_xml = package
                .part(&sheet_part_name)
                .ok_or_else(|| WorkbookPackageError::MissingPart(sheet_part_name.clone()))?;
            let sheet_root =
                XmlElement::parse(sheet_xml).map_err(|source| WorkbookPackageError::Xml {
                    part: sheet_part_name.clone(),
                    source,
                })?;

            let sheet_id = workbook.add_sheet(name.clone())?;
            let sheet_model = workbook.sheet_mut(sheet_id).expect("sheet just added");
            sheet_model.xlsx_sheet_id = xlsx_sheet_id;
            sheet_model.xlsx_rel_id = Some(rel_id.to_string());
            sheet_model.visibility = visibility;
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

    /// Load a workbook package from a file on disk.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn load(path: &Path) -> Result<Self, WorkbookPackageError> {
        let bytes = fs::read(path)?;
        Self::from_bytes(&bytes)
    }

    /// Serialize the workbook package to `.xlsx`/`.xlsm` bytes.
    pub fn write_to_bytes(&mut self) -> Result<Vec<u8>, WorkbookPackageError> {
        // Ensure we have xf indices for any styles referenced by stored cells.
        let style_ids = self
            .workbook
            .sheets
            .iter()
            .flat_map(|sheet| sheet.iter_cells().map(|(_, cell)| cell.style_id));
        self.styles
            .xf_indices_for_style_ids(style_ids, &self.workbook.styles)?;

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

        Ok(self.package.write_to_bytes()?)
    }

    /// Save the workbook package to a file on disk.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save(&mut self, out_path: &Path) -> Result<(), WorkbookPackageError> {
        let bytes = self.write_to_bytes()?;
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
