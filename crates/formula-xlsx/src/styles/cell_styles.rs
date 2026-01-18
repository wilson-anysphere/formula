//! `styles.xml` parsing/writing for cell formats.
//!
//! This module focuses on the `cellXfs` pipeline:
//! - XLSX stores styles as integer indices (`xf` records in `cellXfs`).
//! - `formula-model` stores styles in a deduplicated [`StyleTable`] and cells reference a `style_id`.
//!
//! We parse `styles.xml` into a `StylesPart` that maintains a bidirectional mapping between
//! XLSX `xf` indices and internal `style_id`s. When new styles are introduced, we append new
//! definitions (fonts/fills/borders/numFmts/cellXfs) so existing indices remain stable.

use std::collections::HashMap;

use formula_model::{
    parse_argb_hex_color, Alignment, Border, BorderEdge, BorderStyle, CfStyleOverride, Color, Fill,
    FillPattern, Font, HorizontalAlignment, Protection, Style, StyleTable, VerticalAlignment,
};

use crate::xml::{QName, XmlDomError, XmlElement, XmlNode};

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// Default `styles.xml` payload used when a package omits the styles part.
///
/// This mirrors what Excel generates for a blank workbook and is shared across
/// the `WorkbookPackage` and `XlsxDocument` pipelines.
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
pub enum StylesPartError {
    #[error("allocation failure: {0}")]
    AllocationFailure(&'static str),
    #[error("styles.xml root is not <styleSheet>")]
    InvalidRoot,
    #[error("unknown style_id {0}")]
    UnknownStyleId(u32),
    #[error(transparent)]
    Xml(#[from] XmlDomError),
}

#[derive(Debug, Clone)]
pub struct StylesPart {
    root: XmlElement,

    xf_style_ids: Vec<u32>,
    style_to_xf: HashMap<u32, u32>,

    fonts: Vec<Font>,
    font_index: HashMap<Font, u32>,
    fills: Vec<Fill>,
    fill_index: HashMap<Fill, u32>,
    borders: Vec<Border>,
    border_index: HashMap<Border, u32>,

    num_fmt_by_id: HashMap<u16, String>,
    num_fmt_id_by_code: HashMap<String, u16>,
    next_custom_num_fmt_id: u16,
}

impl StylesPart {
    pub fn parse(bytes: &[u8], style_table: &mut StyleTable) -> Result<Self, StylesPartError> {
        let root = XmlElement::parse(bytes)?;

        if root.name.local != "styleSheet" {
            return Err(StylesPartError::InvalidRoot);
        }

        let num_fmt_by_id = parse_num_fmts(&root);
        let mut num_fmt_id_by_code = HashMap::new();
        let mut max_custom = 163u16;
        for (id, code) in &num_fmt_by_id {
            num_fmt_id_by_code.entry(code.clone()).or_insert(*id);
            max_custom = max_custom.max(*id);
        }

        let next_custom_num_fmt_id = max_custom.saturating_add(1).max(164);

        let fonts = parse_fonts(&root);
        let mut font_index = HashMap::new();
        for (idx, font) in fonts.iter().cloned().enumerate() {
            font_index.entry(font).or_insert(idx as u32);
        }

        let fills = parse_fills(&root);
        let mut fill_index = HashMap::new();
        for (idx, fill) in fills.iter().cloned().enumerate() {
            fill_index.entry(fill).or_insert(idx as u32);
        }

        let borders = parse_borders(&root);
        let mut border_index = HashMap::new();
        for (idx, border) in borders.iter().cloned().enumerate() {
            border_index.entry(border).or_insert(idx as u32);
        }

        let mut xf_style_ids = Vec::new();
        let mut style_to_xf = HashMap::new();

        if let Some(cell_xfs) = root.child("cellXfs") {
            for (xf_idx, xf_el) in cell_xfs.children_by_local("xf").enumerate() {
                let style = parse_xf(xf_el, &fonts, &fills, &borders, &num_fmt_by_id);
                let style_id = style_table.intern(style);
                xf_style_ids.push(style_id);
                style_to_xf.entry(style_id).or_insert(xf_idx as u32);
            }
        } else {
            // Minimal fallback: at least one default xf.
            xf_style_ids.push(style_table.intern(Style::default()));
            style_to_xf.insert(0, 0);
        }

        Ok(Self {
            root,
            xf_style_ids,
            style_to_xf,
            fonts,
            font_index,
            fills,
            fill_index,
            borders,
            border_index,
            num_fmt_by_id,
            num_fmt_id_by_code,
            next_custom_num_fmt_id,
        })
    }

    pub fn parse_or_default(
        bytes: Option<&[u8]>,
        style_table: &mut StyleTable,
    ) -> Result<Self, StylesPartError> {
        match bytes {
            Some(bytes) => Self::parse(bytes, style_table),
            None => Self::parse(DEFAULT_STYLES_XML.as_bytes(), style_table),
        }
    }

    pub fn style_id_for_xf(&self, xf_index: u32) -> u32 {
        self.xf_style_ids
            .get(xf_index as usize)
            .copied()
            .unwrap_or(0)
    }

    /// Return the number of `<xf>` records in `cellXfs`.
    ///
    /// This corresponds to the valid range of worksheet `c/@s` indices.
    pub fn cell_xfs_count(&self) -> usize {
        self.xf_style_ids.len()
    }

    pub fn xf_index_for_style(
        &mut self,
        style_id: u32,
        style_table: &StyleTable,
    ) -> Result<u32, StylesPartError> {
        if let Some(existing) = self.style_to_xf.get(&style_id) {
            return Ok(*existing);
        }

        let style = style_table
            .get(style_id)
            .ok_or(StylesPartError::UnknownStyleId(style_id))?
            .clone();

        let num_fmt_id = self.intern_number_format(style.number_format.as_deref());
        let font_id = self.intern_font(style.font.as_ref());
        let fill_id = self.intern_fill(style.fill.as_ref());
        let border_id = self.intern_border(style.border.as_ref());

        let xf = build_xf_element(
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            style.alignment.as_ref(),
            style.protection.as_ref(),
        );

        let xf_idx = self.append_cell_xf(xf);
        self.xf_style_ids.push(style_id);
        self.style_to_xf.insert(style_id, xf_idx);
        Ok(xf_idx)
    }

    pub fn to_xml_bytes(&self) -> Vec<u8> {
        self.root.to_xml_string().into_bytes()
    }

    /// Parse differential formats (`<dxfs>`) for conditional formatting rules.
    ///
    /// This is a lightweight parser over the already-parsed `styles.xml` DOM held by this
    /// [`StylesPart`]. It avoids re-parsing `styles.xml` when importing conditional formatting.
    pub fn conditional_formatting_dxfs(&self) -> Vec<CfStyleOverride> {
        let Some(dxfs) = self.root.child("dxfs") else {
            return Vec::new();
        };

        dxfs.children_by_local("dxf")
            .map(parse_conditional_formatting_dxf)
            .collect()
    }

    /// Update the `<dxfs>` section (differential formats) for conditional formatting rules.
    ///
    /// Excel stores conditional-formatting style overrides in `styles.xml` under `<dxfs>`,
    /// and worksheet `<cfRule dxfId="...">` attributes reference indices into this list.
    ///
    /// This helper updates the DOM held by this [`StylesPart`] while trying to preserve any
    /// existing `<dxf>` entries that are semantically equivalent to the desired ones. When a
    /// `<dxf>` differs, it is replaced with a minimal representation generated from
    /// [`CfStyleOverride`]. New entries are appended.
    pub fn set_conditional_formatting_dxfs(
        &mut self,
        desired: &[CfStyleOverride],
    ) {
        let dxfs_el = ensure_styles_child(&mut self.root, "dxfs");

        // Positions of existing `<dxf>` element nodes within `dxfs_el.children`.
        let mut existing_indices: Vec<usize> = dxfs_el
            .children
            .iter()
            .enumerate()
            .filter_map(|(idx, child)| match child {
                XmlNode::Element(el) if el.name.local == "dxf" => Some(idx),
                _ => None,
            })
            .collect();

        // Replace or append to reach the desired length.
        for (desired_idx, desired_dxf) in desired.iter().enumerate() {
            if let Some(&child_idx) = existing_indices.get(desired_idx) {
                let should_replace = match dxfs_el.children.get(child_idx) {
                    Some(XmlNode::Element(existing_el)) => {
                        parse_conditional_formatting_dxf(existing_el) != *desired_dxf
                    }
                    _ => true,
                };
                if should_replace {
                    dxfs_el.children[child_idx] =
                        XmlNode::Element(build_conditional_formatting_dxf(desired_dxf));
                }
            } else {
                // Append new `<dxf>` element.
                dxfs_el
                    .children
                    .push(XmlNode::Element(build_conditional_formatting_dxf(desired_dxf)));
            }
        }

        // If the desired list is shorter than the existing one, drop the extra `<dxf>` entries
        // (in reverse order to keep indices stable while removing).
        if existing_indices.len() > desired.len() {
            // Refresh indices because we may have appended above.
            existing_indices = dxfs_el
                .children
                .iter()
                .enumerate()
                .filter_map(|(idx, child)| match child {
                    XmlNode::Element(el) if el.name.local == "dxf" => Some(idx),
                    _ => None,
                })
                .collect();

            for &idx in existing_indices.iter().skip(desired.len()).rev() {
                dxfs_el.children.remove(idx);
            }
        }

        dxfs_el.set_attr("count", desired.len().to_string());
    }

    /// Append additional differential formats (`<dxf>`) to the existing `styles.xml` `<dxfs>` table.
    ///
    /// This is intended for round-trip flows where we want to preserve existing `<dxf>` entries
    /// (including any unknown/unmodeled XML) while still allowing new conditional formatting rules
    /// to reference additional DXFs.
    ///
    /// `dxfs` entries are appended in order. The `<dxfs @count>` attribute is updated to match the
    /// final number of `<dxf>` children.
    pub fn append_conditional_formatting_dxfs(&mut self, dxfs: &[CfStyleOverride]) {
        if dxfs.is_empty() {
            return;
        }

        let dxfs_el = ensure_styles_child(&mut self.root, "dxfs");
        let new_nodes: Vec<XmlNode> = dxfs
            .iter()
            .map(|dxf| XmlNode::Element(build_conditional_formatting_dxf(dxf)))
            .collect();

        // `<dxfs>` may contain an `<extLst>` element which must appear *after* all `<dxf>`
        // children. When appending, preserve this ordering by inserting new `<dxf>` elements
        // immediately before `<extLst>` if present.
        if let Some(ext_lst_idx) = dxfs_el.children.iter().position(|child| {
            matches!(child, XmlNode::Element(el) if el.name.local == "extLst")
        }) {
            dxfs_el
                .children
                .splice(ext_lst_idx..ext_lst_idx, new_nodes);
        } else {
            dxfs_el.children.extend(new_nodes);
        }
        let count = dxfs_el.children_by_local("dxf").count();
        dxfs_el.set_attr("count", count.to_string());
    }

    /// Ensure every `style_id` in `style_ids` has a corresponding `xf` index.
    ///
    /// The returned map can be used to set worksheet `c/@s` attributes.
    ///
    /// `style_ids` are processed in sorted order so new `xf` records are appended
    /// deterministically.
    pub fn xf_indices_for_style_ids(
        &mut self,
        style_ids: impl IntoIterator<Item = u32>,
        style_table: &StyleTable,
    ) -> Result<HashMap<u32, u32>, StylesPartError> {
        let mut ids: Vec<u32> = style_ids.into_iter().collect();
        ids.sort_unstable();
        ids.dedup();

        let mut out = HashMap::new();
        if out.try_reserve(ids.len()).is_err() {
            return Err(StylesPartError::AllocationFailure(
                "xf_indices_for_style_ids output map",
            ));
        }
        for style_id in ids {
            let xf_index = self.xf_index_for_style(style_id, style_table)?;
            out.insert(style_id, xf_index);
        }
        Ok(out)
    }

    /// Resolve a `numFmtId` to an explicit format code when the workbook defines
    /// a custom number format for that id in `styles.xml` (`<numFmts>`).
    ///
    /// Built-in formats are *not* returned here; use
    /// [`formula_format::builtin_format_code`] (or the `__builtin_numFmtId:<id>`
    /// placeholder convention) for those.
    pub fn num_fmt_code_for_id(&self, num_fmt_id: u16) -> Option<&str> {
        self.num_fmt_by_id.get(&num_fmt_id).map(|s| s.as_str())
    }

    fn append_cell_xf(&mut self, xf: XmlElement) -> u32 {
        let cell_xfs = ensure_styles_child(&mut self.root, "cellXfs");
        let count = cell_xfs.children_by_local("xf").count();
        cell_xfs.children.push(XmlNode::Element(xf));
        cell_xfs.set_attr("count", (count + 1).to_string());
        count as u32
    }

    fn intern_font(&mut self, font: Option<&Font>) -> u32 {
        let Some(font) = font.cloned() else {
            return 0;
        };

        if let Some(existing) = self.font_index.get(&font) {
            return *existing;
        }

        let xml_font = build_font_element(&font);
        let fonts_el = ensure_styles_child(&mut self.root, "fonts");
        let idx = fonts_el.children_by_local("font").count();
        fonts_el.children.push(XmlNode::Element(xml_font));
        fonts_el.set_attr("count", (idx + 1).to_string());

        self.fonts.push(font.clone());
        self.font_index.insert(font, idx as u32);
        idx as u32
    }

    fn intern_fill(&mut self, fill: Option<&Fill>) -> u32 {
        let Some(fill) = fill.cloned() else {
            return 0;
        };

        if let Some(existing) = self.fill_index.get(&fill) {
            return *existing;
        }

        let xml_fill = build_fill_element(&fill);
        let fills_el = ensure_styles_child(&mut self.root, "fills");
        let idx = fills_el.children_by_local("fill").count();
        fills_el.children.push(XmlNode::Element(xml_fill));
        fills_el.set_attr("count", (idx + 1).to_string());

        self.fills.push(fill.clone());
        self.fill_index.insert(fill, idx as u32);
        idx as u32
    }

    fn intern_border(&mut self, border: Option<&Border>) -> u32 {
        let Some(border) = border.cloned() else {
            return 0;
        };

        if let Some(existing) = self.border_index.get(&border) {
            return *existing;
        }

        let xml_border = build_border_element(&border);
        let borders_el = ensure_styles_child(&mut self.root, "borders");
        let idx = borders_el.children_by_local("border").count();
        borders_el.children.push(XmlNode::Element(xml_border));
        borders_el.set_attr("count", (idx + 1).to_string());

        self.borders.push(border.clone());
        self.border_index.insert(border, idx as u32);
        idx as u32
    }

    fn intern_number_format(&mut self, fmt: Option<&str>) -> u16 {
        let Some(fmt) = fmt else {
            return 0;
        };

        if let Some(id) = parse_builtin_placeholder(fmt) {
            return id;
        }

        if let Some(id) = builtin_num_fmt_id_for_code(fmt) {
            return id;
        }

        if let Some(id) = self.num_fmt_id_by_code.get(fmt) {
            return *id;
        }

        let id = self.next_custom_num_fmt_id;
        self.next_custom_num_fmt_id = self.next_custom_num_fmt_id.saturating_add(1);
        self.num_fmt_by_id.insert(id, fmt.to_string());
        self.num_fmt_id_by_code.insert(fmt.to_string(), id);

        let num_fmts = ensure_styles_child(&mut self.root, "numFmts");
        let count = num_fmts.children_by_local("numFmt").count();
        num_fmts
            .children
            .push(XmlNode::Element(build_num_fmt_element(id, fmt)));
        num_fmts.set_attr("count", (count + 1).to_string());
        id
    }
}

fn parse_conditional_formatting_dxf(dxf: &XmlElement) -> CfStyleOverride {
    let mut out = CfStyleOverride::default();

    if let Some(font) = dxf.child("font") {
        if let Some(b) = font.children_by_local("b").next() {
            out.bold = Some(parse_dxf_bool(b));
        }
        if let Some(i) = font.children_by_local("i").next() {
            out.italic = Some(parse_dxf_bool(i));
        }
        if let Some(color) = font.child("color") {
            if let Some(rgb) = color.attr("rgb") {
                out.font_color = parse_argb_hex_color(rgb);
            }
        }
    }

    if let Some(fill) = dxf.child("fill") {
        if let Some(pattern_fill) = fill.child("patternFill") {
            if let Some(fg) = pattern_fill.child("fgColor") {
                if let Some(rgb) = fg.attr("rgb") {
                    out.fill = parse_argb_hex_color(rgb);
                }
            }
        }
    }

    out
}

fn parse_dxf_bool(el: &XmlElement) -> bool {
    match el.attr("val") {
        None => true,
        Some(v) => !(v == "0" || v.eq_ignore_ascii_case("false")),
    }
}

fn build_conditional_formatting_dxf(style: &CfStyleOverride) -> XmlElement {
    let mut dxf = empty_element("dxf");

    if let Some(fill_color) = style.fill {
        let mut fill = empty_element("fill");
        let mut pattern_fill = empty_element("patternFill");
        pattern_fill.set_attr("patternType", "solid");
        pattern_fill
            .children
            .push(XmlNode::Element(build_color_element("fgColor", fill_color)));
        let mut bg = empty_element("bgColor");
        bg.set_attr("indexed", "64");
        pattern_fill.children.push(XmlNode::Element(bg));
        fill.children.push(XmlNode::Element(pattern_fill));
        dxf.children.push(XmlNode::Element(fill));
    }

    let has_font = style.bold.is_some() || style.italic.is_some() || style.font_color.is_some();
    if has_font {
        let mut font = empty_element("font");
        if let Some(bold) = style.bold {
            if bold {
                font.children.push(XmlNode::Element(empty_element("b")));
            } else {
                let mut b = empty_element("b");
                b.set_attr("val", "0");
                font.children.push(XmlNode::Element(b));
            }
        }
        if let Some(italic) = style.italic {
            if italic {
                font.children.push(XmlNode::Element(empty_element("i")));
            } else {
                let mut i = empty_element("i");
                i.set_attr("val", "0");
                font.children.push(XmlNode::Element(i));
            }
        }
        if let Some(color) = style.font_color {
            font.children
                .push(XmlNode::Element(build_color_element("color", color)));
        }
        dxf.children.push(XmlNode::Element(font));
    }
    dxf
}
fn parse_num_fmts(root: &XmlElement) -> HashMap<u16, String> {
    let mut out = HashMap::new();
    let Some(num_fmts) = root.child("numFmts") else {
        return out;
    };

    for num_fmt in num_fmts.children_by_local("numFmt") {
        let id = num_fmt.attr("numFmtId").and_then(|v| v.parse::<u16>().ok());
        let code = num_fmt.attr("formatCode").map(|s| s.to_string());
        if let (Some(id), Some(code)) = (id, code) {
            out.insert(id, code);
        }
    }
    out
}

fn parse_fonts(root: &XmlElement) -> Vec<Font> {
    let Some(fonts) = root.child("fonts") else {
        return vec![Font::default()];
    };

    let mut parsed: Vec<Font> = fonts.children_by_local("font").map(parse_font).collect();
    if parsed.is_empty() {
        parsed.push(Font::default());
        return parsed;
    }

    // Normalize each font entry against index 0 so internal styles only store the deltas.
    let base = parsed[0].clone();
    for font in &mut parsed {
        normalize_font(font, &base);
    }

    parsed
}

fn parse_font(el: &XmlElement) -> Font {
    let name = el
        .child("name")
        .and_then(|n| n.attr("val"))
        .map(|s| s.to_string());
    let size_100pt = el
        .child("sz")
        .and_then(|sz| sz.attr("val"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| (v * 100.0).round() as u16);

    let bold = el.child("b").is_some();
    let italic = el.child("i").is_some();
    let underline = el
        .child("u")
        .is_some_and(|u| u.attr("val").map(|v| v != "none").unwrap_or(true));
    let strike = el.child("strike").is_some();

    let color = el.child("color").and_then(parse_color);

    Font {
        name,
        size_100pt,
        bold,
        italic,
        underline,
        strike,
        color,
    }
}

fn normalize_font(font: &mut Font, base: &Font) {
    if font.name == base.name {
        font.name = None;
    }
    if font.size_100pt == base.size_100pt {
        font.size_100pt = None;
    }
    if font.color == base.color {
        font.color = None;
    }

    // Boolean flags are stored as deltas implicitly (base is expected to be false).
}

fn parse_fills(root: &XmlElement) -> Vec<Fill> {
    let Some(fills) = root.child("fills") else {
        return vec![
            Fill::default(),
            Fill {
                pattern: FillPattern::Gray125,
                ..Fill::default()
            },
        ];
    };

    let mut out = Vec::new();
    for fill in fills.children_by_local("fill") {
        out.push(parse_fill(fill));
    }
    if out.is_empty() {
        out.push(Fill::default());
    }
    out
}

fn parse_fill(el: &XmlElement) -> Fill {
    let Some(pattern_fill) = el.child("patternFill") else {
        return Fill::default();
    };

    let pattern = match pattern_fill.attr("patternType").unwrap_or("none") {
        "none" => FillPattern::None,
        "gray125" => FillPattern::Gray125,
        "solid" => FillPattern::Solid,
        other => FillPattern::Other(other.to_string()),
    };

    let fg_color = pattern_fill.child("fgColor").and_then(parse_color);
    let bg_color = pattern_fill.child("bgColor").and_then(parse_color);

    Fill {
        pattern,
        fg_color,
        bg_color,
    }
}

fn parse_borders(root: &XmlElement) -> Vec<Border> {
    let Some(borders) = root.child("borders") else {
        return vec![Border::default()];
    };

    let mut out = Vec::new();
    for border in borders.children_by_local("border") {
        out.push(parse_border(border));
    }
    if out.is_empty() {
        out.push(Border::default());
    }
    out
}

fn parse_border(el: &XmlElement) -> Border {
    let left = parse_border_edge(el.child("left"));
    let right = parse_border_edge(el.child("right"));
    let top = parse_border_edge(el.child("top"));
    let bottom = parse_border_edge(el.child("bottom"));
    let diagonal = parse_border_edge(el.child("diagonal"));

    let diagonal_up = el
        .attr("diagonalUp")
        .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true"));
    let diagonal_down = el
        .attr("diagonalDown")
        .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true"));

    Border {
        left,
        right,
        top,
        bottom,
        diagonal,
        diagonal_up,
        diagonal_down,
    }
}

fn parse_border_edge(edge: Option<&XmlElement>) -> BorderEdge {
    let Some(edge) = edge else {
        return BorderEdge::default();
    };

    let style = match edge.attr("style").unwrap_or("none") {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "double" => BorderStyle::Double,
        _ => BorderStyle::None,
    };
    let color = edge.child("color").and_then(parse_color);

    BorderEdge { style, color }
}

fn parse_xf(
    xf: &XmlElement,
    fonts: &[Font],
    fills: &[Fill],
    borders: &[Border],
    num_fmts: &HashMap<u16, String>,
) -> Style {
    let font_id = xf
        .attr("fontId")
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);
    let fill_id = xf
        .attr("fillId")
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);
    let border_id = xf
        .attr("borderId")
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);

    let font = if font_id == 0 {
        None
    } else {
        fonts
            .get(font_id)
            .cloned()
            .filter(|f| f != &Font::default())
    };
    let fill = fills
        .get(fill_id)
        .cloned()
        .filter(|f| !is_default_fill(f, fill_id));
    let border = borders
        .get(border_id)
        .cloned()
        .filter(|b| !is_default_border(b, border_id));

    let alignment = xf.child("alignment").and_then(parse_alignment);
    let protection = xf.child("protection").and_then(parse_protection);

    let num_fmt_id = xf
        .attr("numFmtId")
        .and_then(|v| v.parse::<u16>().ok())
        .unwrap_or(0);
    let number_format = if num_fmt_id == 0 {
        None
    } else if let Some(code) = num_fmts.get(&num_fmt_id) {
        Some(code.clone())
    } else if let Some(code) = builtin_num_fmt_code(num_fmt_id) {
        Some(code.to_string())
    } else {
        Some(format!(
            "{}{num_fmt_id}",
            formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX
        ))
    };

    Style {
        font,
        fill,
        border,
        alignment,
        protection,
        number_format,
    }
}

fn is_default_fill(fill: &Fill, fill_id: usize) -> bool {
    fill_id == 0 && matches!(fill.pattern, FillPattern::None)
}

fn is_default_border(border: &Border, border_id: usize) -> bool {
    border_id == 0 && border == &Border::default()
}

fn parse_alignment(el: &XmlElement) -> Option<Alignment> {
    let horizontal = el.attr("horizontal").and_then(parse_horizontal_alignment);
    let vertical = el.attr("vertical").and_then(parse_vertical_alignment);
    let wrap_text = el
        .attr("wrapText")
        .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true"));
    let rotation = el.attr("textRotation").and_then(|v| v.parse::<i16>().ok());
    let indent = el.attr("indent").and_then(|v| v.parse::<u16>().ok());

    let alignment = Alignment {
        horizontal,
        vertical,
        wrap_text,
        rotation,
        indent,
    };

    if alignment == Alignment::default() {
        None
    } else {
        Some(alignment)
    }
}

fn parse_protection(el: &XmlElement) -> Option<Protection> {
    let locked = el.attr("locked").map(|v| v != "0").unwrap_or(true);
    let hidden = el.attr("hidden").is_some_and(|v| v != "0");

    let protection = Protection { locked, hidden };
    if protection == Protection::default() {
        None
    } else {
        Some(protection)
    }
}

fn parse_horizontal_alignment(value: &str) -> Option<HorizontalAlignment> {
    match value {
        "general" => Some(HorizontalAlignment::General),
        "left" => Some(HorizontalAlignment::Left),
        "center" => Some(HorizontalAlignment::Center),
        "right" => Some(HorizontalAlignment::Right),
        "fill" => Some(HorizontalAlignment::Fill),
        "justify" => Some(HorizontalAlignment::Justify),
        _ => None,
    }
}

fn parse_vertical_alignment(value: &str) -> Option<VerticalAlignment> {
    match value {
        "top" => Some(VerticalAlignment::Top),
        "center" => Some(VerticalAlignment::Center),
        "bottom" => Some(VerticalAlignment::Bottom),
        _ => None,
    }
}

fn parse_color(el: &XmlElement) -> Option<Color> {
    if el
        .attr("auto")
        .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    {
        return Some(Color::Auto);
    }

    if let Some(rgb) = el.attr("rgb") {
        return parse_argb(rgb).map(Color::Argb);
    }

    if let Some(theme) = el.attr("theme").and_then(|v| v.parse::<u16>().ok()) {
        let tint = el
            .attr("tint")
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| (v.clamp(-1.0, 1.0) * 1000.0).round() as i16);
        return Some(Color::Theme { theme, tint });
    }

    if let Some(indexed) = el.attr("indexed").and_then(|v| v.parse::<u16>().ok()) {
        return Some(Color::Indexed(indexed));
    }

    None
}

fn parse_argb(value: &str) -> Option<u32> {
    let hex = value.trim();
    let hex = hex.strip_prefix('#').unwrap_or(hex);
    if hex.len() == 8 {
        u32::from_str_radix(hex, 16).ok()
    } else if hex.len() == 6 {
        u32::from_str_radix(hex, 16)
            .ok()
            .map(|rgb| 0xFF00_0000 | rgb)
    } else {
        None
    }
}

fn build_xf_element(
    num_fmt_id: u16,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    alignment: Option<&Alignment>,
    protection: Option<&Protection>,
) -> XmlElement {
    let mut xf = XmlElement {
        name: QName {
            ns: Some(NS_MAIN.to_string()),
            local: "xf".to_string(),
        },
        attrs: Default::default(),
        children: Vec::new(),
    };

    xf.set_attr("numFmtId", num_fmt_id.to_string());
    xf.set_attr("fontId", font_id.to_string());
    xf.set_attr("fillId", fill_id.to_string());
    xf.set_attr("borderId", border_id.to_string());
    xf.set_attr("xfId", "0");

    if num_fmt_id != 0 {
        xf.set_attr("applyNumberFormat", "1");
    }
    if font_id != 0 {
        xf.set_attr("applyFont", "1");
    }
    if fill_id != 0 {
        xf.set_attr("applyFill", "1");
    }
    if border_id != 0 {
        xf.set_attr("applyBorder", "1");
    }

    if let Some(alignment) = alignment {
        xf.set_attr("applyAlignment", "1");
        xf.children
            .push(XmlNode::Element(build_alignment_element(alignment)));
    }

    if let Some(protection) = protection {
        xf.set_attr("applyProtection", "1");
        xf.children
            .push(XmlNode::Element(build_protection_element(protection)));
    }

    xf
}

fn build_font_element(font: &Font) -> XmlElement {
    let mut el = XmlElement {
        name: QName {
            ns: Some(NS_MAIN.to_string()),
            local: "font".to_string(),
        },
        attrs: Default::default(),
        children: Vec::new(),
    };

    if font.bold {
        el.children.push(XmlNode::Element(empty_element("b")));
    }
    if font.italic {
        el.children.push(XmlNode::Element(empty_element("i")));
    }
    if font.underline {
        el.children.push(XmlNode::Element(empty_element("u")));
    }
    if font.strike {
        el.children.push(XmlNode::Element(empty_element("strike")));
    }

    if let Some(size) = font.size_100pt {
        let mut sz = empty_element("sz");
        sz.set_attr("val", format!("{:.2}", (size as f32) / 100.0));
        el.children.push(XmlNode::Element(sz));
    }

    if let Some(color) = font.color {
        el.children
            .push(XmlNode::Element(build_color_element("color", color)));
    }

    if let Some(name) = &font.name {
        let mut n = empty_element("name");
        n.set_attr("val", name.clone());
        el.children.push(XmlNode::Element(n));
    }

    el
}

fn build_fill_element(fill: &Fill) -> XmlElement {
    let mut el = XmlElement {
        name: QName {
            ns: Some(NS_MAIN.to_string()),
            local: "fill".to_string(),
        },
        attrs: Default::default(),
        children: Vec::new(),
    };

    let mut pattern_fill = empty_element("patternFill");
    let pattern_type = match &fill.pattern {
        FillPattern::None => "none".to_string(),
        FillPattern::Gray125 => "gray125".to_string(),
        FillPattern::Solid => "solid".to_string(),
        FillPattern::Other(value) => value.clone(),
    };
    pattern_fill.set_attr("patternType", pattern_type);

    if let Some(color) = fill.fg_color {
        pattern_fill
            .children
            .push(XmlNode::Element(build_color_element("fgColor", color)));
    }
    if let Some(color) = fill.bg_color {
        pattern_fill
            .children
            .push(XmlNode::Element(build_color_element("bgColor", color)));
    }

    el.children.push(XmlNode::Element(pattern_fill));
    el
}

fn build_border_element(border: &Border) -> XmlElement {
    let mut el = XmlElement {
        name: QName {
            ns: Some(NS_MAIN.to_string()),
            local: "border".to_string(),
        },
        attrs: Default::default(),
        children: Vec::new(),
    };

    if border.diagonal_up {
        el.set_attr("diagonalUp", "1");
    }
    if border.diagonal_down {
        el.set_attr("diagonalDown", "1");
    }

    el.children.push(XmlNode::Element(build_border_edge_element(
        "left",
        &border.left,
    )));
    el.children.push(XmlNode::Element(build_border_edge_element(
        "right",
        &border.right,
    )));
    el.children.push(XmlNode::Element(build_border_edge_element(
        "top",
        &border.top,
    )));
    el.children.push(XmlNode::Element(build_border_edge_element(
        "bottom",
        &border.bottom,
    )));
    el.children.push(XmlNode::Element(build_border_edge_element(
        "diagonal",
        &border.diagonal,
    )));

    el
}

fn build_border_edge_element(name: &str, edge: &BorderEdge) -> XmlElement {
    let mut el = empty_element(name);
    let style = match edge.style {
        BorderStyle::None => None,
        BorderStyle::Thin => Some("thin"),
        BorderStyle::Medium => Some("medium"),
        BorderStyle::Thick => Some("thick"),
        BorderStyle::Dashed => Some("dashed"),
        BorderStyle::Dotted => Some("dotted"),
        BorderStyle::Double => Some("double"),
    };
    if let Some(style) = style {
        el.set_attr("style", style);
    }
    if let Some(color) = edge.color {
        el.children
            .push(XmlNode::Element(build_color_element("color", color)));
    }
    el
}

fn build_alignment_element(alignment: &Alignment) -> XmlElement {
    let mut el = empty_element("alignment");
    if let Some(horizontal) = alignment.horizontal {
        el.set_attr("horizontal", format_horizontal_alignment(horizontal));
    }
    if let Some(vertical) = alignment.vertical {
        el.set_attr("vertical", format_vertical_alignment(vertical));
    }
    if alignment.wrap_text {
        el.set_attr("wrapText", "1");
    }
    if let Some(rotation) = alignment.rotation {
        el.set_attr("textRotation", rotation.to_string());
    }
    if let Some(indent) = alignment.indent {
        el.set_attr("indent", indent.to_string());
    }
    el
}

fn build_protection_element(protection: &Protection) -> XmlElement {
    let mut el = empty_element("protection");
    if !protection.locked {
        el.set_attr("locked", "0");
    }
    if protection.hidden {
        el.set_attr("hidden", "1");
    }
    el
}

fn build_color_element(name: &str, color: Color) -> XmlElement {
    let mut el = empty_element(name);
    match color {
        Color::Argb(argb) => el.set_attr("rgb", format!("{:08X}", argb)),
        Color::Theme { theme, tint } => {
            el.set_attr("theme", theme.to_string());
            if let Some(tint) = tint {
                el.set_attr("tint", format!("{:.3}", (tint as f64) / 1000.0));
            }
        }
        Color::Indexed(index) => el.set_attr("indexed", index.to_string()),
        Color::Auto => el.set_attr("auto", "1"),
    }
    el
}

fn build_num_fmt_element(id: u16, code: &str) -> XmlElement {
    let mut el = empty_element("numFmt");
    el.set_attr("numFmtId", id.to_string());
    el.set_attr("formatCode", code.to_string());
    el
}

fn builtin_num_fmt_code(id: u16) -> Option<&'static str> {
    match id {
        // Only a small subset of built-ins are expanded to format codes during
        // parse; the rest are preserved via `__builtin_numFmtId:<id>` so we can
        // round-trip IDs exactly even when multiple ids share the same code.
        7 | 9 | 14 => formula_format::builtin_format_code(id),
        _ => None,
    }
}

fn builtin_num_fmt_id_for_code(code: &str) -> Option<u16> {
    // Some writers emit simplified forms of Excel built-ins (especially currency).
    // Accept these aliases so we can re-emit built-in ids instead of introducing
    // custom `numFmtId` entries.
    if code == "$#,##0.00" {
        return Some(7);
    }

    formula_format::builtin_format_id(code)
}

fn parse_builtin_placeholder(code: &str) -> Option<u16> {
    let rest = code.strip_prefix(formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX)?;
    rest.parse::<u16>().ok()
}

fn empty_element(local: &str) -> XmlElement {
    XmlElement {
        name: QName {
            ns: Some(NS_MAIN.to_string()),
            local: local.to_string(),
        },
        attrs: Default::default(),
        children: Vec::new(),
    }
}

fn format_horizontal_alignment(alignment: HorizontalAlignment) -> &'static str {
    match alignment {
        HorizontalAlignment::General => "general",
        HorizontalAlignment::Left => "left",
        HorizontalAlignment::Center => "center",
        HorizontalAlignment::Right => "right",
        HorizontalAlignment::Fill => "fill",
        HorizontalAlignment::Justify => "justify",
    }
}

fn format_vertical_alignment(alignment: VerticalAlignment) -> &'static str {
    match alignment {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Bottom => "bottom",
    }
}

fn ensure_styles_child<'a>(root: &'a mut XmlElement, local: &str) -> &'a mut XmlElement {
    let mut idx = root
        .children
        .iter()
        .position(|child| matches!(child, XmlNode::Element(el) if el.name.local == local))
        .unwrap_or_else(|| {
            let idx = insertion_index(root, local);
            root.children.insert(idx, XmlNode::Element(empty_element(local)));
            idx
        });

    if idx >= root.children.len() {
        debug_assert!(
            false,
            "styles child insertion produced out-of-bounds index; falling back to append"
        );
        idx = root.children.len();
        root.children.push(XmlNode::Element(empty_element(local)));
    }

    let node = &mut root.children[idx];
    loop {
        match node {
            XmlNode::Element(el) => return el,
            _ => {
                debug_assert!(
                    false,
                    "styles child index should point at an XmlNode::Element; repairing in place"
                );
                *node = XmlNode::Element(empty_element(local));
            }
        }
    }
}

fn insertion_index(root: &XmlElement, local: &str) -> usize {
    let order = [
        "numFmts",
        "fonts",
        "fills",
        "borders",
        "cellStyleXfs",
        "cellXfs",
        "cellStyles",
        "dxfs",
        "tableStyles",
        "extLst",
    ];

    let Some(target_pos) = order.iter().position(|name| name == &local) else {
        return root.children.len();
    };

    for (idx, child) in root.children.iter().enumerate() {
        let XmlNode::Element(el) = child else {
            continue;
        };
        if let Some(pos) = order.iter().position(|name| name == &el.name.local) {
            if pos > target_pos {
                return idx;
            }
        }
    }

    root.children.len()
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    fn fixture(path: &str) -> String {
        let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join(path);
        dir.to_string_lossy().to_string()
    }

    fn load_fixture_styles_xml(path: &str) -> Vec<u8> {
        let bytes = std::fs::read(fixture(path)).expect("fixture exists");
        let mut zip =
            zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid fixture zip");
        let mut file = zip.by_name("xl/styles.xml").expect("styles.xml exists");
        let mut out = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut out).expect("read styles.xml");
        out
    }

    #[test]
    fn conditional_formatting_dxf_parses_bold_italic_val_zero_as_false() {
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dxfs count="1">
    <dxf>
      <font>
        <b val="0"/>
        <i val="false"/>
      </font>
    </dxf>
  </dxfs>
</styleSheet>
"#;

        let mut style_table = StyleTable::new();
        let styles_part = StylesPart::parse(styles_xml.as_bytes(), &mut style_table).unwrap();
        let dxfs = styles_part.conditional_formatting_dxfs();

        assert_eq!(dxfs.len(), 1);
        assert_eq!(dxfs[0].bold, Some(false));
        assert_eq!(dxfs[0].italic, Some(false));
    }

    #[test]
    fn writes_conditional_formatting_dxfs_into_default_styles() {
        let mut style_table = StyleTable::new();
        let mut part = StylesPart::parse_or_default(None, &mut style_table).unwrap();

        let dxfs = vec![
            CfStyleOverride {
                fill: Some(Color::new_argb(0xFFFF0000)),
                font_color: Some(Color::new_argb(0xFFFFFFFF)),
                bold: Some(true),
                italic: Some(false),
            },
            CfStyleOverride {
                fill: Some(Color::new_argb(0xFF00FF00)),
                font_color: None,
                bold: None,
                italic: None,
            },
            CfStyleOverride {
                fill: None,
                font_color: Some(Color::new_argb(0xFF0000FF)),
                bold: Some(false),
                italic: Some(true),
            },
        ];

        part.set_conditional_formatting_dxfs(&dxfs);
        assert_eq!(part.conditional_formatting_dxfs(), dxfs);

        let xml = String::from_utf8(part.to_xml_bytes()).unwrap();
        let doc = Document::parse(&xml).unwrap();
        let main_ns = NS_MAIN;
        let dxfs_node = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "dxfs" && n.tag_name().namespace() == Some(main_ns))
            .expect("dxfs element present");
        assert_eq!(dxfs_node.attribute("count"), Some("3"));
        let dxf_nodes: Vec<_> = dxfs_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "dxf")
            .collect();
        assert_eq!(dxf_nodes.len(), 3);

        // dxf[0]: fill + font, with italic disabled.
        let fill = dxf_nodes[0]
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "fill")
            .expect("fill present");
        let pattern_fill = fill
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "patternFill")
            .unwrap();
        assert_eq!(pattern_fill.attribute("patternType"), Some("solid"));
        let fg = pattern_fill
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "fgColor")
            .unwrap();
        assert_eq!(fg.attribute("rgb"), Some("FFFF0000"));
        let bg = pattern_fill
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "bgColor")
            .unwrap();
        assert_eq!(bg.attribute("indexed"), Some("64"));

        let font = dxf_nodes[0]
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "font")
            .expect("font present");
        assert!(font
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "b"));
        let i = font
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "i")
            .expect("italic element present");
        assert_eq!(i.attribute("val"), Some("0"));
        let color = font
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "color")
            .unwrap();
        assert_eq!(color.attribute("rgb"), Some("FFFFFFFF"));

        // dxf[1]: fill only, no font element.
        assert!(dxf_nodes[1]
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "fill"));
        assert!(!dxf_nodes[1]
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "font"));

        // dxf[2]: font only, no fill element.
        assert!(!dxf_nodes[2]
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "fill"));
        assert!(dxf_nodes[2]
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "font"));
    }

    #[test]
    fn inserts_dxfs_when_missing_and_keeps_order() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#;

        let mut style_table = StyleTable::new();
        let mut part = StylesPart::parse(xml.as_bytes(), &mut style_table).unwrap();
        part.set_conditional_formatting_dxfs(&[]);

        let xml = String::from_utf8(part.to_xml_bytes()).unwrap();
        let doc = Document::parse(&xml).unwrap();
        let root = doc.root_element();
        let children: Vec<_> = root
            .children()
            .filter(|n| n.is_element())
            .map(|n| n.tag_name().name())
            .collect();

        let dxfs_idx = children
            .iter()
            .position(|name| *name == "dxfs")
            .expect("dxfs inserted");
        let table_styles_idx = children
            .iter()
            .position(|name| *name == "tableStyles")
            .expect("tableStyles present");
        assert!(dxfs_idx < table_styles_idx, "dxfs should appear before tableStyles");
    }

    #[test]
    fn rewriting_fixture_dxfs_preserves_semantics_and_other_children() {
        let styles_xml = load_fixture_styles_xml("conditional_formatting_2007.xlsx");
        let mut style_table = StyleTable::new();
        let mut part = StylesPart::parse(&styles_xml, &mut style_table).unwrap();
        let original_root = part.root.clone();
        let dxfs = part.conditional_formatting_dxfs();

        part.set_conditional_formatting_dxfs(&dxfs);
        assert_eq!(part.conditional_formatting_dxfs(), dxfs);

        let strip_dxfs = |root: &XmlElement| {
            root.children
                .iter()
                .filter(|child| !matches!(child, XmlNode::Element(el) if el.name.local == "dxfs"))
                .cloned()
                .collect::<Vec<_>>()
        };
        assert_eq!(strip_dxfs(&part.root), strip_dxfs(&original_root));
    }
}
