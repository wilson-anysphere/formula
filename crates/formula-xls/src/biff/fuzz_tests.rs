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
        // Keep these fuzz-style tests deterministic in CI so failures are reproducible and don't
        // depend on a random per-run seed.
        rng_seed: proptest::test_runner::RngSeed::Fixed(0),
        // Avoid writing `proptest-regressions/` files during CI or local runs of this fuzz-style
        // suite. If a panic is found, we want the minimized input printed in the test failure and
        // for the fix to land quickly, rather than persisting stateful regression artifacts.
        failure_persistence: None,
        .. ProptestConfig::default()
    })]

    #[test]
    fn biff_parsers_are_panic_free_on_arbitrary_input(buf in proptest::collection::vec(any::<u8>(), 0..=MAX_INPUT_LEN)) {
        let start = offset_in_bounds(&buf);

        // Physical BIFF record iteration.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let mut iter = records::BiffRecordIter::from_offset(&buf, start).expect("offset derived to be in-bounds");
                let mut prev_end = start;
                while let Some(next) = iter.next() {
                    match next {
                        Ok(record) => {
                            // Basic iterator invariants: records are contiguous, in-bounds slices.
                            assert!(record.offset >= prev_end, "record offsets should be monotonic");
                            let end = record
                                .offset
                                .checked_add(4)
                                .and_then(|v| v.checked_add(record.data.len()))
                                .expect("overflow computing record end");
                            assert!(end <= buf.len(), "record extends past end of input");
                            prev_end = end;
                        }
                        Err(_) => {
                            // An error terminates iteration.
                            assert!(iter.next().is_none());
                            break;
                        }
                    }
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
                while let Some(next) = iter.next() {
                    match next {
                        Ok(record) => {
                            // Fragment boundaries must be consistent with the combined payload.
                            let total: usize = record.fragment_sizes.iter().copied().sum();
                            assert_eq!(total, record.data.len(), "fragment sizes must sum to combined data length");

                            let frag_lens: Vec<usize> = record.fragments().map(|f| f.len()).collect();
                            assert_eq!(frag_lens, record.fragment_sizes, "fragment iterator must match fragment_sizes");

                            if record.fragment_sizes.len() <= 1 {
                                assert!(!record.is_continued());
                            } else {
                                assert!(record.is_continued());
                            }
                        }
                        Err(_) => {
                            // Malformed physical record terminates the logical iterator too.
                            break;
                        }
                    }
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
                let parsed =
                    defined_names::parse_biff_defined_names(&buf, BiffVersion::Biff8, CODEPAGE_1252, sheet_names)
                        .expect("defined name parsing is best-effort and should not hard-fail");
                for name in parsed.names {
                    assert!(!name.name.is_empty(), "defined name should not be empty");
                    assert!(!name.name.contains('\0'), "defined name should have embedded NULs stripped");
                    if let Some(comment) = name.comment {
                        assert!(!comment.contains('\0'), "defined name comment should have embedded NULs stripped");
                    }
                }
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
                for note in parsed.notes {
                    assert!(!note.author.contains('\0'), "NOTE author should have embedded NULs stripped");
                    assert!(!note.text.contains('\0'), "NOTE text should have embedded NULs stripped");
                }
            }))
            .is_ok(),
            "comments::parse_biff_sheet_notes panicked"
        );

        // Worksheet HLINK parser.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let parsed = sheet::parse_biff_sheet_hyperlinks(&buf, 0, CODEPAGE_1252)
                    .expect("offset 0 should always be in-bounds");
                for link in parsed.hyperlinks {
                    assert!(link.range.start.row <= link.range.end.row);
                    assert!(link.range.start.col <= link.range.end.col);
                }
            }))
            .is_ok(),
            "sheet::parse_biff_sheet_hyperlinks panicked"
        );

        // AutoFilter `_FilterDatabase` NAME parser.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let parsed =
                    autofilter::parse_biff_filter_database_ranges(&buf, BiffVersion::Biff8, CODEPAGE_1252, None)
                        .expect("autofilter parsing is best-effort and should not hard-fail");
                for (_sheet_idx, range) in parsed.by_sheet {
                    assert!(range.start.row <= range.end.row);
                    assert!(range.start.col <= range.end.col);
                }
            }))
            .is_ok(),
            "autofilter::parse_biff_filter_database_ranges panicked"
        );
    }
}
