use proptest::prelude::*;

use super::{autofilter, comments, defined_names, globals, records, sheet, strings, BiffVersion};

const MAX_INPUT_LEN: usize = 64 * 1024;
const CODEPAGE_1252: u16 = 1252;
const MAX_SHEET_NOTE_WARNINGS: usize = 20;

fn offset_in_bounds(buf: &[u8]) -> usize {
    // Derive a "random-ish" offset from the first few bytes while always staying in-bounds so we
    // can exercise `*_iter::from_offset` on arbitrary inputs without relying on a separate proptest
    // strategy.
    if buf.is_empty() {
        return 0;
    }

    let b0 = buf[0] as usize;
    let b1 = buf.get(1).copied().unwrap_or(0) as usize;
    let candidate = (b0 << 8) | b1;
    candidate % (buf.len() + 1)
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 256,
        .. ProptestConfig::default()
    })]

    #[test]
    fn biff_parsers_are_panic_free_on_arbitrary_input(buf in proptest::collection::vec(any::<u8>(), 0..=MAX_INPUT_LEN)) {
        let start = offset_in_bounds(&buf);

        // Physical BIFF record iteration.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let mut iter = records::BiffRecordIter::from_offset(&buf, start).expect("offset derived to be in-bounds");
                while let Some(_next) = iter.next() {
                    // Intentionally discard output; the property is "never panic".
                }
            }))
            .is_ok(),
            "records::BiffRecordIter panicked"
        );

        // Logical BIFF record iteration (CONTINUE coalescing).
        let allows_continuation: fn(u16) -> bool = |_| true;
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let mut iter = records::LogicalBiffRecordIter::from_offset(&buf, start, allows_continuation)
                    .expect("offset derived to be in-bounds");
                while let Some(_next) = iter.next() {
                    // Discard.
                }
            }))
            .is_ok(),
            "records::LogicalBiffRecordIter panicked"
        );

        // Best-effort BIFF8 Unicode string decoding.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = strings::parse_biff8_unicode_string_best_effort(&buf, CODEPAGE_1252);
            }))
            .is_ok(),
            "strings::parse_biff8_unicode_string_best_effort panicked"
        );

        // Workbook globals parser (used by `.xls` importer).
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = globals::parse_biff_workbook_globals(&buf, BiffVersion::Biff8, CODEPAGE_1252);
            }))
            .is_ok(),
            "globals::parse_biff_workbook_globals panicked"
        );

        // Defined names parser.
        let sheet_names: &[String] = &[];
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = defined_names::parse_biff_defined_names(&buf, BiffVersion::Biff8, CODEPAGE_1252, sheet_names);
            }))
            .is_ok(),
            "defined_names::parse_biff_defined_names panicked"
        );

        // Worksheet NOTE/OBJ/TXO comment ("notes") parser. This parser should be panic-free and
        // should keep warnings bounded per sheet to avoid unbounded memory growth on corrupt inputs.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let parsed = comments::parse_biff_sheet_notes(&buf, 0, BiffVersion::Biff8, CODEPAGE_1252)
                    .expect("offset 0 should always be in-bounds");
                assert!(
                    parsed.warnings.len() <= MAX_SHEET_NOTE_WARNINGS,
                    "warnings should be bounded (len={})",
                    parsed.warnings.len()
                );
            }))
            .is_ok(),
            "comments::parse_biff_sheet_notes panicked"
        );

        // Worksheet HLINK parser.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = sheet::parse_biff_sheet_hyperlinks(&buf, 0, CODEPAGE_1252)
                    .expect("offset 0 should always be in-bounds");
            }))
            .is_ok(),
            "sheet::parse_biff_sheet_hyperlinks panicked"
        );

        // AutoFilter `_FilterDatabase` NAME parser.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = autofilter::parse_biff_filter_database_ranges(&buf, BiffVersion::Biff8, CODEPAGE_1252, None);
            }))
            .is_ok(),
            "autofilter::parse_biff_filter_database_ranges panicked"
        );
    }
}

