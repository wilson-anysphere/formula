use std::collections::HashMap;

use formula_model::{CfStyleOverride, Worksheet, WorksheetId};

/// Aggregated conditional formatting differential formats (`<dxfs>`) for an XLSX workbook.
///
/// In OOXML, the `styles.xml` part contains a single global `<dxfs>` table and
/// worksheet `<cfRule dxfId="...">` attributes index into that table.
///
/// `formula_model::Worksheet`, however, stores `conditional_formatting_dxfs` per-sheet.
/// Writers should use this helper to:
/// - Build a deterministic global dxfs vector (deduping identical entries across sheets).
/// - Remap each worksheet-local `dxf_id` to the global index when emitting `<cfRule dxfId>`.
///
/// Deduplication is stable: the first time a particular [`CfStyleOverride`] is seen (in sheet
/// iteration order, then local vector order) it "wins" and determines the global index.
#[derive(Clone, Debug, Default)]
pub struct ConditionalFormattingDxfAggregation {
    /// Global, deduplicated `<dxfs>` entries to emit into `xl/styles.xml`.
    pub global_dxfs: Vec<CfStyleOverride>,
    /// Per-sheet mapping from local `dxf_id` (index into `Worksheet::conditional_formatting_dxfs`)
    /// to global `dxf_id` (index into [`Self::global_dxfs`]).
    pub local_to_global_by_sheet: HashMap<WorksheetId, Vec<u32>>,
}

impl ConditionalFormattingDxfAggregation {
    /// Build a global `<dxfs>` table and per-sheet local->global mappings.
    pub fn from_worksheets<'a>(worksheets: impl IntoIterator<Item = &'a Worksheet>) -> Self {
        let mut global_dxfs: Vec<CfStyleOverride> = Vec::new();
        let mut local_to_global_by_sheet: HashMap<WorksheetId, Vec<u32>> = HashMap::new();

        for sheet in worksheets {
            let mut mapping: Vec<u32> = Vec::new();
            let _ = mapping.try_reserve_exact(sheet.conditional_formatting_dxfs.len());

            for local in &sheet.conditional_formatting_dxfs {
                let existing = global_dxfs.iter().position(|g| g == local);
                let global_idx = match existing {
                    Some(idx) => idx as u32,
                    None => {
                        let idx = global_dxfs.len() as u32;
                        global_dxfs.push(local.clone());
                        idx
                    }
                };
                mapping.push(global_idx);
            }

            local_to_global_by_sheet.insert(sheet.id, mapping);
        }

        Self {
            global_dxfs,
            local_to_global_by_sheet,
        }
    }

    /// Build a global `<dxfs>` table with a pre-existing base table.
    ///
    /// This is primarily intended for round-trip writers (e.g. `XlsxDocument`) that are editing an
    /// existing workbook package:
    /// - The workbook already has a `styles.xml` `<dxfs>` table with stable indices referenced by
    ///   existing worksheet `<cfRule dxfId="...">` attributes.
    /// - We want to preserve those existing indices/entries, and only append new dxfs when the
    ///   in-memory model introduces additional differential formats.
    ///
    /// The returned `global_dxfs` starts with `base_global_dxfs` (in the provided order), followed
    /// by any new dxfs encountered while iterating worksheets (deduped by exact equality).
    pub fn from_worksheets_with_base_global_dxfs<'a>(
        worksheets: impl IntoIterator<Item = &'a Worksheet>,
        base_global_dxfs: &[CfStyleOverride],
    ) -> Self {
        let mut global_dxfs: Vec<CfStyleOverride> = base_global_dxfs.to_vec();
        let mut local_to_global_by_sheet: HashMap<WorksheetId, Vec<u32>> = HashMap::new();

        for sheet in worksheets {
            let mut mapping: Vec<u32> = Vec::new();
            let _ = mapping.try_reserve_exact(sheet.conditional_formatting_dxfs.len());

            for (local_idx, local) in sheet.conditional_formatting_dxfs.iter().enumerate() {
                // Fast-path: if the worksheet-local dxfs vector is aligned with the current global
                // table, preserve the index. This is common when a workbook was loaded from an
                // existing `styles.xml` and each sheet stores the global dxfs vector as-is.
                let global_idx = if local_idx < global_dxfs.len() && global_dxfs[local_idx] == *local
                {
                    local_idx as u32
                } else if let Some(idx) = global_dxfs.iter().position(|g| g == local) {
                    idx as u32
                } else {
                    let idx = global_dxfs.len() as u32;
                    global_dxfs.push(local.clone());
                    idx
                };
                mapping.push(global_idx);
            }

            local_to_global_by_sheet.insert(sheet.id, mapping);
        }

        Self {
            global_dxfs,
            local_to_global_by_sheet,
        }
    }

    /// Map a worksheet-local `dxf_id` to the global `dxf_id`.
    ///
    /// Returns `None` when:
    /// - `dxf_id` is `None`, or
    /// - `dxf_id` is out of bounds for the sheet's local `conditional_formatting_dxfs` vector, or
    /// - the sheet has no mapping (e.g. it was not included in the aggregation).
    ///
    /// Writers should treat `None` as "emit no `dxfId` attribute" (best-effort).
    pub fn map_rule_dxf_id(&self, sheet_id: WorksheetId, dxf_id: Option<u32>) -> Option<u32> {
        let local = dxf_id?;
        let mapping = self.local_to_global_by_sheet.get(&sheet_id)?;
        mapping.get(local as usize).copied()
    }
}
