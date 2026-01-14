// Dedicated test target for workbook information functions (SHEET/SHEETS/etc).
//
// The bulk of these tests live under `tests/functions/` and are also included in the aggregated
// `financial` integration test crate. Exposing them as a standalone target allows running:
// `cargo test -p formula-engine --test workbook_info`.
#[path = "functions/workbook_info.rs"]
mod workbook_info;
