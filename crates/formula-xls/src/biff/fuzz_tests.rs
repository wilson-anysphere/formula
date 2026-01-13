use proptest::prelude::*;

use super::{
    autofilter, comments, defined_names, encryption, globals, records, sheet, strings, BiffVersion,
};

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

        // Best-effort BIFF substream iteration.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let mut iter =
                    records::BestEffortSubstreamIter::from_offset(&buf, start).expect("offset derived to be in-bounds");
                while let Some(record) = iter.next() {
                    // Basic invariants: returned slices must be in-bounds.
                    let end = record
                        .offset
                        .checked_add(4)
                        .and_then(|v| v.checked_add(record.data.len()))
                        .expect("overflow computing record end");
                    assert!(end <= buf.len(), "record extends past end of input");
                }
            }))
            .is_ok(),
            "records::BestEffortSubstreamIter panicked"
        );

        // Encryption preflight scan helper.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = records::workbook_globals_has_filepass_record(&buf);
            }))
            .is_ok(),
            "records::workbook_globals_has_filepass_record panicked"
        );

        // FILEPASS parser (encryption classifier).
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let payload_len = buf.len().min(encryption::MAX_FILEPASS_PAYLOAD_BYTES);
                let _ = encryption::parse_filepass_record(BiffVersion::Biff8, &buf[..payload_len]);
            }))
            .is_ok(),
            "encryption::parse_filepass_record panicked"
        );

        // Workbook decrypt preflight helper (should fail gracefully on arbitrary inputs).
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let mut workbook = buf.clone();
                let _ = encryption::decrypt_workbook_stream(&mut workbook, "pw");
            }))
            .is_ok(),
            "encryption::decrypt_workbook_stream panicked"
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

                // Exercise the strict string parsers as well; they should return an error on
                // malformed payloads rather than panic.
                let _ = strings::parse_biff5_short_string(&buf, CODEPAGE_1252);
                let _ = strings::parse_biff8_short_string(&buf, CODEPAGE_1252);
                let _ = strings::parse_biff8_unicode_string(&buf, CODEPAGE_1252);
                let _ = strings::parse_biff_short_string(&buf, BiffVersion::Biff5, CODEPAGE_1252);
                let _ = strings::parse_biff_short_string(&buf, BiffVersion::Biff8, CODEPAGE_1252);
                let _ = strings::parse_biff5_short_string_best_effort(&buf, CODEPAGE_1252);

                // Continued string parser (uses the fragment-aware cursor).
                let fragments: [&[u8]; 1] = [&buf];
                let _ = strings::parse_biff8_unicode_string_continued(&fragments, 0, CODEPAGE_1252);
            }))
            .is_ok(),
            "strings::parse_biff8_unicode_string_best_effort panicked"
        );

        // Workbook globals parser (used by `.xls` importer).
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = globals::parse_biff_workbook_globals(&buf, BiffVersion::Biff8, CODEPAGE_1252);
                let _ = globals::parse_biff_workbook_globals(&buf, BiffVersion::Biff5, CODEPAGE_1252);
            }))
            .is_ok(),
            "globals::parse_biff_workbook_globals panicked"
        );

        // Additional workbook-global helpers used by the importer.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = globals::parse_biff_codepage(&buf);
                let sheets = globals::parse_biff_bound_sheets(&buf, BiffVersion::Biff8, CODEPAGE_1252)
                    .expect("offset 0 should always be in-bounds");
                for sheet in sheets {
                    assert!(!sheet.name.contains('\0'), "BoundSheet name should have embedded NULs stripped");
                }
            }))
            .is_ok(),
            "globals::parse_biff_codepage/parse_biff_bound_sheets panicked"
        );

        // Defined names parser.
        let sheet_names: &[String] = &[];
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let parsed =
                    defined_names::parse_biff_defined_names(&buf, BiffVersion::Biff8, CODEPAGE_1252, sheet_names)
                        .expect("defined name parsing is best-effort and should not hard-fail");
                let _ = defined_names::parse_biff_defined_names(&buf, BiffVersion::Biff5, CODEPAGE_1252, sheet_names);
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

        // Worksheet metadata parsers used by the importer.
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let protection = sheet::parse_biff_sheet_protection(&buf, 0).expect("offset 0 should always be in-bounds");
                // Basic invariant: password hash is present iff non-zero (enforced by parser).
                if let Some(hash) = protection.protection.password_hash {
                    assert_ne!(hash, 0);
                }

                let view = sheet::parse_biff_sheet_view_state(&buf, 0).expect("offset 0 should always be in-bounds");
                if let Some(zoom) = view.zoom {
                    assert!(zoom.is_finite() && zoom > 0.0);
                }

                let row_col = sheet::parse_biff_sheet_row_col_properties(&buf, 0, CODEPAGE_1252)
                    .expect("offset 0 should always be in-bounds");
                if let Some(range) = row_col.auto_filter_range {
                    assert!(range.start.row <= range.end.row);
                    assert!(range.start.col <= range.end.col);
                }
                if let Some(sort) = row_col.sort_state {
                    for cond in sort.conditions {
                        assert!(cond.range.start.row <= cond.range.end.row);
                        assert!(cond.range.start.col <= cond.range.end.col);
                    }
                }
                for (_row, props) in row_col.rows {
                    if let Some(h) = props.height {
                        assert!(h.is_finite() && h > 0.0);
                    }
                    assert!(props.outline_level <= 7);
                }
                for (_col, props) in row_col.cols {
                    if let Some(w) = props.width {
                        assert!(w.is_finite() && w > 0.0);
                    }
                    assert!(props.outline_level <= 7);
                }

                let merged = sheet::parse_biff_sheet_merged_cells(&buf, 0)
                    .expect("offset 0 should always be in-bounds");
                for range in merged.ranges {
                    assert!(range.start.row <= range.end.row);
                    assert!(range.start.col <= range.end.col);
                }

                // Manual page breaks.
                let _ = sheet::parse_biff_sheet_manual_page_breaks(&buf, 0)
                    .expect("offset 0 should always be in-bounds");

                // Print settings helper (page setup + margins + manual page breaks).
                let print = super::parse_biff_sheet_print_settings(&buf, 0)
                    .expect("offset 0 should always be in-bounds");
                if let Some(page_setup) = print.page_setup.as_ref() {
                    let margins = &page_setup.margins;
                    assert!(margins.left.is_finite());
                    assert!(margins.right.is_finite());
                    assert!(margins.top.is_finite());
                    assert!(margins.bottom.is_finite());
                    assert!(margins.header.is_finite());
                    assert!(margins.footer.is_finite());
                }

                let xfs = sheet::parse_biff_sheet_cell_xf_indices_filtered(&buf, 0, None)
                    .expect("offset 0 should always be in-bounds");
                for (cell, _xf) in xfs {
                    assert!(cell.row < formula_model::EXCEL_MAX_ROWS);
                    assert!(cell.col < formula_model::EXCEL_MAX_COLS);
                }
            }))
            .is_ok(),
            "sheet metadata parsers panicked"
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
                let _ = autofilter::parse_biff_filter_database_ranges(&buf, BiffVersion::Biff5, CODEPAGE_1252, None);
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
