use std::collections::{BTreeMap, BTreeSet};

use crate::preserve::opc_graph::collect_transitive_related_parts;
use crate::XlsxPackage;

impl XlsxPackage {
    /// Extract `xl/media/*` parts referenced by the workbook's RichData (`xl/richData`) graph.
    ///
    /// Excel stores "images in cell" and other RichData values using a set of parts under
    /// `xl/richData/` whose `.rels` can reference `xl/media/*` payloads.
    ///
    /// This function is best-effort:
    /// - Missing/malformed `.rels` parts are ignored.
    /// - `TargetMode="External"` relationships are ignored.
    /// - Missing targets are ignored.
    pub fn rich_data_media_parts(&self) -> BTreeMap<String, Vec<u8>> {
        // Treat every RichData XML part as a traversal root. This catches the common
        // `xl/richData/richValueRel*.xml` relationship sources and is resilient to new
        // RichData subgraphs.
        let root_parts: BTreeSet<String> = self
            .part_names()
            .filter(|name| name.starts_with("xl/richData/"))
            .filter(|name| !name.contains("/_rels/"))
            .filter(|name| name.ends_with(".xml"))
            .map(|name| name.to_string())
            .collect();

        if root_parts.is_empty() {
            return BTreeMap::new();
        }

        let related = match collect_transitive_related_parts(self, root_parts.into_iter()) {
            Ok(parts) => parts,
            // The traversal is intentionally best-effort; treat any error as an empty result.
            Err(_) => BTreeMap::new(),
        };

        related
            .into_iter()
            .filter(|(name, _)| name.starts_with("xl/media/"))
            .collect()
    }
}

