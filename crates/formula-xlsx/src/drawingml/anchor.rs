use formula_model::charts::ChartAnchor;
use formula_model::drawings::{Anchor, AnchorPoint, CellOffset, EmuSize};
use formula_model::CellRef;
use roxmltree::Node;

/// Parse a DrawingML anchor (`xdr:absoluteAnchor`, `xdr:oneCellAnchor`, `xdr:twoCellAnchor`).
///
/// This is shared between the legacy chart extractor (`ChartAnchor`) and the newer drawing
/// pipeline (`drawings::Anchor`) to ensure they stay in lockstep.
///
/// Notes:
/// - `<xdr:colOff>` / `<xdr:rowOff>` are optional in the wild; when missing they default to 0.
/// - Whitespace around numeric values is tolerated.
pub(crate) fn parse_anchor(anchor: &Node<'_, '_>) -> Option<Anchor> {
    match anchor.tag_name().name() {
        "absoluteAnchor" => {
            let pos = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "pos")?;
            let ext = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "ext")?;

            Some(Anchor::Absolute {
                pos: CellOffset::new(parse_attr_i64(pos, "x")?, parse_attr_i64(pos, "y")?),
                ext: EmuSize::new(parse_attr_i64(ext, "cx")?, parse_attr_i64(ext, "cy")?),
            })
        }
        "oneCellAnchor" => {
            let from = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "from")?;
            let ext = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "ext")?;

            Some(Anchor::OneCell {
                from: parse_anchor_point(from)?,
                ext: EmuSize::new(parse_attr_i64(ext, "cx")?, parse_attr_i64(ext, "cy")?),
            })
        }
        "twoCellAnchor" => {
            let from = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "from")?;
            let to = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "to")?;

            Some(Anchor::TwoCell {
                from: parse_anchor_point(from)?,
                to: parse_anchor_point(to)?,
            })
        }
        _ => None,
    }
}

/// Return the `<xdr:*Anchor>` nodes inside a worksheet drawing (`<xdr:wsDr>`), handling
/// `mc:AlternateContent` wrappers.
///
/// Some producers wrap anchors in `mc:AlternateContent` (with both `mc:Choice` and
/// `mc:Fallback` branches). A naive `.descendants()` search will find anchors in *both*
/// branches, resulting in duplicate chart refs / drawing objects.
///
/// This helper:
/// - walks only direct children of `<xdr:wsDr>`,
/// - treats `mc:AlternateContent` as transparent, selecting the **first** `mc:Choice`
///   branch that contains any anchor nodes (falling back to `mc:Fallback` when no choice
///   contains anchors),
/// - and returns the matching anchor nodes in document order.
pub(crate) fn wsdr_anchor_nodes<'a, 'input>(wsdr: Node<'a, 'input>) -> Vec<Node<'a, 'input>> {
    fn is_anchor_node(node: Node<'_, '_>) -> bool {
        node.is_element()
            && matches!(
                node.tag_name().name(),
                "oneCellAnchor" | "twoCellAnchor" | "absoluteAnchor"
            )
    }

    fn anchors_in_branch<'a, 'input>(branch: Node<'a, 'input>) -> Vec<Node<'a, 'input>> {
        branch
            .descendants()
            .filter(|n| is_anchor_node(*n))
            .collect()
    }

    let mut out = Vec::new();
    for child in wsdr.children().filter(|n| n.is_element()) {
        if is_anchor_node(child) {
            out.push(child);
            continue;
        }

        if child.tag_name().name() != "AlternateContent" {
            continue;
        }

        // Prefer the first Choice branch that contains anchors.
        let mut selected: Option<Vec<Node<'a, 'input>>> = None;
        for choice in child
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "Choice")
        {
            let anchors = anchors_in_branch(choice);
            if !anchors.is_empty() {
                selected = Some(anchors);
                break;
            }
        }

        // Fall back when no Choice branch contains anchors.
        if selected.is_none() {
            for fallback in child
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "Fallback")
            {
                let anchors = anchors_in_branch(fallback);
                if !anchors.is_empty() {
                    selected = Some(anchors);
                    break;
                }
            }
        }

        if let Some(anchors) = selected {
            out.extend(anchors);
        }
    }

    out
}

/// Flattens `mc:AlternateContent` wrappers by selecting a single branch.
///
/// This mirrors the chartSpace parser's heuristic:
/// - Prefer the first `mc:Choice` branch that contains any node matching `selector`
///   (searching within that branch).
/// - Otherwise, prefer the first `mc:Fallback` branch that contains any node matching `selector`.
/// - Otherwise, fall back to the first non-empty branch.
///
/// The returned nodes are the selected branch's element children (recursively flattened).
pub(crate) fn flatten_alternate_content<'a, 'input>(
    node: Node<'a, 'input>,
    selector: fn(Node<'a, 'input>) -> bool,
) -> Vec<Node<'a, 'input>> {
    if node.tag_name().name() != "AlternateContent" {
        return vec![node];
    }

    let mut first_choice_children: Option<Vec<Node<'a, 'input>>> = None;
    for choice in node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "Choice")
    {
        let children: Vec<_> = choice
            .children()
            .filter(|n| n.is_element())
            .flat_map(|n| flatten_alternate_content(n, selector))
            .collect();
        if first_choice_children.is_none() && !children.is_empty() {
            first_choice_children = Some(children.clone());
        }
        if choice.descendants().any(selector) {
            return children;
        }
    }

    let mut first_fallback_children: Option<Vec<Node<'a, 'input>>> = None;
    for fallback in node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "Fallback")
    {
        let children: Vec<_> = fallback
            .children()
            .filter(|n| n.is_element())
            .flat_map(|n| flatten_alternate_content(n, selector))
            .collect();
        if first_fallback_children.is_none() && !children.is_empty() {
            first_fallback_children = Some(children.clone());
        }
        if fallback.descendants().any(selector) {
            return children;
        }
    }

    if let Some(children) = first_choice_children {
        return children;
    }
    if let Some(children) = first_fallback_children {
        return children;
    }

    // Unknown structure: treat AlternateContent as transparent and just emit its
    // direct element children.
    node.children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, selector))
        .collect()
}

/// Returns the direct element children of `node`, flattening `mc:AlternateContent` wrappers.
pub(crate) fn element_children_selecting_alternate_content<'a, 'input>(
    node: Node<'a, 'input>,
    selector: fn(Node<'a, 'input>) -> bool,
) -> Vec<Node<'a, 'input>> {
    node.children()
        .filter(|n| n.is_element())
        .flat_map(|n| flatten_alternate_content(n, selector))
        .collect()
}

/// Returns all descendants of `node` matching `desired`, traversing `mc:AlternateContent` wrappers
/// by selecting a single branch using `selector`.
pub(crate) fn descendants_selecting_alternate_content<'a, 'input>(
    node: Node<'a, 'input>,
    selector: fn(Node<'a, 'input>) -> bool,
    desired: fn(Node<'a, 'input>) -> bool,
) -> Vec<Node<'a, 'input>> {
    fn walk<'a, 'input>(
        node: Node<'a, 'input>,
        selector: fn(Node<'a, 'input>) -> bool,
        desired: fn(Node<'a, 'input>) -> bool,
        out: &mut Vec<Node<'a, 'input>>,
    ) {
        for child in node.children().filter(|n| n.is_element()) {
            for node in flatten_alternate_content(child, selector) {
                if desired(node) {
                    out.push(node);
                }
                walk(node, selector, desired, out);
            }
        }
    }

    let mut out = Vec::new();
    walk(node, selector, desired, &mut out);
    out
}

pub(crate) fn anchor_to_chart_anchor(anchor: Anchor) -> ChartAnchor {
    match anchor {
        Anchor::TwoCell { from, to } => ChartAnchor::TwoCell {
            from_col: from.cell.col,
            from_row: from.cell.row,
            from_col_off_emu: from.offset.x_emu,
            from_row_off_emu: from.offset.y_emu,
            to_col: to.cell.col,
            to_row: to.cell.row,
            to_col_off_emu: to.offset.x_emu,
            to_row_off_emu: to.offset.y_emu,
        },
        Anchor::OneCell { from, ext } => ChartAnchor::OneCell {
            from_col: from.cell.col,
            from_row: from.cell.row,
            from_col_off_emu: from.offset.x_emu,
            from_row_off_emu: from.offset.y_emu,
            cx_emu: ext.cx,
            cy_emu: ext.cy,
        },
        Anchor::Absolute { pos, ext } => ChartAnchor::Absolute {
            x_emu: pos.x_emu,
            y_emu: pos.y_emu,
            cx_emu: ext.cx,
            cy_emu: ext.cy,
        },
    }
}

pub(crate) fn parse_anchor_point(node: Node<'_, '_>) -> Option<AnchorPoint> {
    let col: u32 = child_text(node, "col")?.trim().parse().ok()?;
    let row: u32 = child_text(node, "row")?.trim().parse().ok()?;

    let col_off: i64 = parse_optional_i64_child(node, "colOff").unwrap_or(0);
    let row_off: i64 = parse_optional_i64_child(node, "rowOff").unwrap_or(0);

    Some(AnchorPoint::new(
        CellRef::new(row, col),
        CellOffset::new(col_off, row_off),
    ))
}

fn parse_attr_i64(node: Node<'_, '_>, attr: &str) -> Option<i64> {
    node.attribute(attr)?.trim().parse::<i64>().ok()
}

fn parse_optional_i64_child(node: Node<'_, '_>, tag: &str) -> Option<i64> {
    child_text(node, tag)?.trim().parse::<i64>().ok()
}

fn child_text<'a>(node: Node<'a, 'a>, tag: &str) -> Option<&'a str> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == tag)
        .and_then(|n| n.text())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::drawingml::charts::extract_chart_object_refs;
    use crate::drawingml::extract_chart_refs;
    use pretty_assertions::assert_eq;

    #[test]
    fn two_cell_anchor_parses_identically_in_both_extractors() {
        let xml = r#"
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col> 1 </xdr:col>
      <xdr:colOff> 2 </xdr:colOff>
      <xdr:row> 3 </xdr:row>
      <xdr:rowOff> 4 </xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col> 5 </xdr:col>
      <xdr:colOff> 6 </xdr:colOff>
      <xdr:row> 7 </xdr:row>
      <xdr:rowOff> 8 </xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"#;

        let legacy = extract_chart_refs(xml.as_bytes(), "drawing.xml").unwrap();
        let modern = extract_chart_object_refs(xml.as_bytes(), "drawing.xml").unwrap();

        assert_eq!(legacy.len(), 1);
        assert_eq!(modern.len(), 1);
        assert_eq!(legacy[0].rel_id, modern[0].rel_id);
        assert_eq!(legacy[0].anchor, anchor_to_chart_anchor(modern[0].anchor));
    }

    #[test]
    fn one_cell_anchor_parses_identically_in_both_extractors() {
        let xml = r#"
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col> 1 </xdr:col>
      <xdr:colOff> 2 </xdr:colOff>
      <xdr:row> 3 </xdr:row>
      <xdr:rowOff> 4 </xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="300" cy="400"/>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>
"#;

        let legacy = extract_chart_refs(xml.as_bytes(), "drawing.xml").unwrap();
        let modern = extract_chart_object_refs(xml.as_bytes(), "drawing.xml").unwrap();

        assert_eq!(legacy.len(), 1);
        assert_eq!(modern.len(), 1);
        assert_eq!(legacy[0].rel_id, modern[0].rel_id);
        assert_eq!(legacy[0].anchor, anchor_to_chart_anchor(modern[0].anchor));
    }

    #[test]
    fn absolute_anchor_parses_identically_in_both_extractors() {
        let xml = r#"
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="100" y="200"/>
    <xdr:ext cx="300" cy="400"/>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>
"#;

        let legacy = extract_chart_refs(xml.as_bytes(), "drawing.xml").unwrap();
        let modern = extract_chart_object_refs(xml.as_bytes(), "drawing.xml").unwrap();

        assert_eq!(legacy.len(), 1);
        assert_eq!(modern.len(), 1);
        assert_eq!(legacy[0].rel_id, modern[0].rel_id);
        assert_eq!(legacy[0].anchor, anchor_to_chart_anchor(modern[0].anchor));
    }
}
