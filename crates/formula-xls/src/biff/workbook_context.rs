//! Shared workbook-global metadata needed to decode BIFF8 `rgce` formula token streams.
//!
//! Both workbook-global defined names (`NAME` records) and worksheet formulas can reference the
//! same workbook-global tables:
//! - `SUPBOOK` / `EXTERNNAME` (external workbook and name metadata)
//! - `EXTERNSHEET` (sheet mapping for 3D references)
//! - `NAME` record order (for `PtgName` indices)
//!
//! This module centralizes best-effort construction of that context so all BIFF8 formula decoding
//! paths resolve `PtgName` consistently.

#![allow(dead_code)]

use super::{defined_names, externsheet, records, rgce, supbook, BiffVersion};

// Record ids used by workbook-global defined name parsing.
// See [MS-XLS] sections:
// - NAME: 2.4.150
const RECORD_NAME: u16 = 0x0018;

/// Workbook-global tables required to decode BIFF8 `rgce` streams.
#[derive(Debug, Clone, Default)]
pub(crate) struct BiffWorkbookContextTables {
    pub(super) codepage: u16,
    pub(super) supbooks: Vec<supbook::SupBookInfo>,
    pub(super) externsheet: Vec<externsheet::ExternSheetEntry>,
    /// NAME metadata in workbook `NAME` record order, including placeholders for unparseable NAME
    /// records so `PtgName` indices remain stable.
    pub(super) defined_names: Vec<rgce::DefinedNameMeta>,
    /// Parsed NAME records in record order, or `None` when a NAME record could not be parsed.
    pub(super) name_records: Vec<Option<defined_names::RawDefinedName>>,
    pub(super) warnings: Vec<String>,
}

impl BiffWorkbookContextTables {
    pub(crate) fn rgce_decode_context<'a>(
        &'a self,
        sheet_names: &'a [String],
    ) -> rgce::RgceDecodeContext<'a> {
        rgce::RgceDecodeContext {
            codepage: self.codepage,
            sheet_names,
            externsheet: &self.externsheet,
            supbooks: &self.supbooks,
            defined_names: &self.defined_names,
        }
    }

    pub(crate) fn drain_warnings(&mut self) -> Vec<String> {
        std::mem::take(&mut self.warnings)
    }
}

/// Best-effort parse of workbook-global rgce decode context (SUPBOOK, EXTERNSHEET, and NAME order
/// metadata).
///
/// This helper never hard-fails on malformed workbook-global tables: parse errors are surfaced as
/// warnings and the returned context contains empty/partial tables.
pub(crate) fn build_biff_workbook_context_tables(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
    sheet_names: &[String],
) -> BiffWorkbookContextTables {
    let mut out = BiffWorkbookContextTables {
        codepage,
        ..Default::default()
    };

    if biff != BiffVersion::Biff8 {
        out.warnings
            .push("BIFF rgce decoding context currently supports BIFF8 only".to_string());
        return out;
    }

    // SUPBOOK/EXTERNNAME and EXTERNSHEET are both best-effort parsers that surface issues as
    // warnings and return whatever metadata could be recovered.
    let supbook::SupBookTable {
        supbooks,
        warnings: supbook_warnings,
    } = supbook::parse_biff8_supbook_table(workbook_stream, codepage);
    out.supbooks = supbooks;
    out.warnings.extend(supbook_warnings);

    let externsheet::ExternSheetTable {
        entries: externsheet_entries,
        warnings,
    } = externsheet::parse_biff_externsheet(workbook_stream, biff, codepage);
    out.externsheet = externsheet_entries;
    out.warnings.extend(warnings);

    // Scan workbook NAME records to build the ordered `PtgName` metadata table.
    let allows_continuation = |id: u16| id == RECORD_NAME;
    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                out.warnings.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        // Stop at the next substream BOF; workbook globals start at offset 0.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_NAME => {
                match defined_names::parse_biff8_name_record(&record, codepage, sheet_names) {
                    Ok(raw) => {
                        out.defined_names.push(rgce::DefinedNameMeta {
                            name: raw.name.clone(),
                            scope_sheet: raw.scope_sheet,
                        });
                        out.name_records.push(Some(raw));
                    }
                    Err(err) => {
                        out.warnings
                            .push(format!("failed to parse NAME record: {err}"));
                        out.defined_names.push(rgce::DefinedNameMeta {
                            name: "#NAME?".to_string(),
                            scope_sheet: None,
                        });
                        out.name_records.push(None);
                    }
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    fn xl_unicode_string_no_cch_compressed(s: &str) -> Vec<u8> {
        // BIFF8 XLUnicodeStringNoCch: [flags: u8][chars]. Emit compressed (8-bit) strings.
        let mut out = Vec::<u8>::new();
        out.push(0); // flags (fHighByte=0)
        out.extend_from_slice(s.as_bytes());
        out
    }

    #[test]
    fn preserves_name_record_order_and_inserts_placeholders() {
        // First NAME record: malformed (truncated description string) so it should produce a
        // placeholder meta in position 0.
        let bad_name = "BadDesc";
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt 1

        let mut bad_header = Vec::new();
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        bad_header.push(0); // chKey
        bad_header.push(bad_name.len() as u8); // cch
        bad_header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        bad_header.extend_from_slice(&0u16.to_le_bytes()); // itab
        bad_header.push(0); // cchCustMenu
        bad_header.push(5); // cchDescription (claims 5 chars, but we truncate below)
        bad_header.push(0); // cchHelpTopic
        bad_header.push(0); // cchStatusText

        let bad_name_str = xl_unicode_string_no_cch_compressed(bad_name);
        // Truncated description: flags + only 2 bytes ("AB"), but header says 5 chars.
        let bad_desc_partial: Vec<u8> = [vec![0u8], b"AB".to_vec()].concat();

        let bad_record_payload =
            [bad_header, bad_name_str, rgce.clone(), bad_desc_partial].concat();

        // Second NAME record: valid.
        let good_name = "Good";
        let mut good_header = Vec::new();
        good_header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        good_header.push(0); // chKey
        good_header.push(good_name.len() as u8); // cch
        good_header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        good_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        good_header.extend_from_slice(&0u16.to_le_bytes()); // itab
        good_header.extend_from_slice(&[0, 0, 0, 0]); // no optional strings

        let good_name_str = xl_unicode_string_no_cch_compressed(good_name);
        let good_record_payload = [good_header, good_name_str, rgce].concat();

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &bad_record_payload),
            record(RECORD_NAME, &good_record_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let ctx = build_biff_workbook_context_tables(&stream, BiffVersion::Biff8, 1252, &[]);
        assert_eq!(ctx.defined_names.len(), 2);
        assert_eq!(ctx.defined_names[0].name, "#NAME?");
        assert_eq!(ctx.defined_names[0].scope_sheet, None);
        assert_eq!(ctx.defined_names[1].name, good_name);
        assert_eq!(ctx.defined_names[1].scope_sheet, None);

        assert_eq!(ctx.name_records.len(), 2);
        assert!(ctx.name_records[0].is_none());
        assert!(ctx.name_records[1].is_some());
    }

    #[test]
    fn warns_on_malformed_supbook_and_externsheet_but_returns_usable_context() {
        // BOF record, followed by a truncated record header/payload that triggers a BIFF iterator
        // error during SUPBOOK/EXTERNSHEET scans.
        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x01AEu16.to_le_bytes()); // SUPBOOK id
        truncated.extend_from_slice(&4u16.to_le_bytes()); // declares 4 bytes...
        truncated.extend_from_slice(&[0xAA, 0xBB]); // ...but only provides 2.

        let stream = [record(records::RECORD_BOF_BIFF8, &[0u8; 16]), truncated].concat();

        let ctx_tables = build_biff_workbook_context_tables(&stream, BiffVersion::Biff8, 1252, &[]);
        assert!(
            ctx_tables
                .warnings
                .iter()
                .any(|w| w.contains("SUPBOOK") && w.contains("malformed BIFF record")),
            "expected SUPBOOK warning, got {:?}",
            ctx_tables.warnings
        );
        assert!(
            ctx_tables
                .warnings
                .iter()
                .any(|w| w.contains("EXTERNSHEET") && w.contains("malformed BIFF record")),
            "expected EXTERNSHEET warning, got {:?}",
            ctx_tables.warnings
        );

        // Context should still be usable for formulas that don't need those tables.
        let ctx = ctx_tables.rgce_decode_context(&[]);
        let decoded = rgce::decode_biff8_rgce(&[0x1E, 0x01, 0x00], &ctx); // PtgInt 1
        assert_eq!(decoded.text, "1");
    }
}
