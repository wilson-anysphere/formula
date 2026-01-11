use std::collections::{BTreeMap, HashSet, VecDeque};

use crate::path::{rels_for_part, resolve_target};
use crate::relationships::parse_relationships;
use crate::workbook::ChartExtractionError;
use crate::XlsxPackage;

/// Collect a transitive closure of OPC parts reachable from `root_parts` via `.rels` files.
///
/// The closure includes:
/// - Each root part itself (when present in the package).
/// - The part's corresponding `.rels` (when present).
/// - All internal relationship targets that exist in the package, recursively.
pub(crate) fn collect_transitive_related_parts(
    pkg: &XlsxPackage,
    root_parts: impl IntoIterator<Item = String>,
) -> Result<BTreeMap<String, Vec<u8>>, ChartExtractionError> {
    let mut out: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let mut visited: HashSet<String> = HashSet::new();
    let mut queue: VecDeque<String> = root_parts.into_iter().collect();

    while let Some(part_name) = queue.pop_front() {
        if !visited.insert(part_name.clone()) {
            continue;
        }

        let Some(part_bytes) = pkg.part(&part_name) else {
            continue;
        };
        out.insert(part_name.clone(), part_bytes.to_vec());

        let rels_part_name = rels_for_part(&part_name);
        let Some(rels_bytes) = pkg.part(&rels_part_name) else {
            continue;
        };
        out.insert(rels_part_name.clone(), rels_bytes.to_vec());

        for rel in parse_relationships(rels_bytes, &rels_part_name)? {
            let target_part = resolve_target(&part_name, &rel.target);
            if pkg.part(&target_part).is_some() {
                queue.push_back(target_part);
            }
        }
    }

    Ok(out)
}

