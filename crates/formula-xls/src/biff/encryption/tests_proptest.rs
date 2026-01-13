use proptest::prelude::*;

use super::*;

const MAX_RECORD_PAYLOAD_LEN: usize = 2048;
const MAX_GLOBAL_RECORDS: usize = 8;
const MAX_SHEET_RECORDS: usize = 6;
const MAX_FILEPASS_PAYLOAD_LEN: usize = 64;

const XOR_BASE_KEY: u8 = 0xA5;

#[derive(Debug, Clone)]
struct RecordSpec {
    record_id: u16,
    payload: Vec<u8>,
}

#[derive(Debug, Clone)]
struct GeneratedStream {
    bytes: Vec<u8>,
    filepass_present: bool,
    // (data_start, len) ranges for every record payload after FILEPASS. Used to apply the test
    // cipher independently of the decryptor's record-walker implementation.
    payload_ranges_after_filepass: Vec<(usize, usize)>,
}

fn xor_keystream_byte(block: u32, block_offset: usize) -> u8 {
    // Mix both low and high bits of the 0..1023 block offset so we actually exercise the block
    // boundary bookkeeping.
    XOR_BASE_KEY ^ (block as u8) ^ (block_offset as u8) ^ ((block_offset >> 8) as u8)
}

fn xor_cipher(block: u32, block_offset: usize, buf: &mut [u8]) -> Result<(), DecryptError> {
    // The record-walking decryptor is expected to split chunks at 1024-byte block boundaries.
    assert!(
        block_offset + buf.len() <= RC4_BLOCK_SIZE,
        "cipher called with chunk that crosses a block boundary (block_offset={}, len={})",
        block_offset,
        buf.len()
    );

    for (idx, b) in buf.iter_mut().enumerate() {
        let offset = block_offset + idx;
        *b ^= xor_keystream_byte(block, offset);
    }
    Ok(())
}

fn apply_xor_spec(stream: &mut [u8], ranges: &[(usize, usize)]) {
    let mut pos = 0usize;
    for &(start, len) in ranges {
        for i in 0..len {
            let block = (pos / RC4_BLOCK_SIZE) as u32;
            let block_offset = pos % RC4_BLOCK_SIZE;
            stream[start + i] ^= xor_keystream_byte(block, block_offset);
            pos += 1;
        }
    }
}

fn push_record(out: &mut Vec<u8>, record_id: u16, payload: &[u8]) -> usize {
    let offset = out.len();
    out.extend_from_slice(&record_id.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    out.extend_from_slice(payload);
    offset
}

fn collect_physical_meta(stream: &[u8]) -> Vec<(usize, u16, usize)> {
    let mut out = Vec::new();
    let mut iter = records::BiffRecordIter::from_offset(stream, 0).expect("offset 0 in-bounds");
    while let Some(next) = iter.next() {
        let record = next.expect("generated stream should be structurally valid");
        out.push((record.offset, record.record_id, record.data.len()));
    }
    out
}

fn assert_physical_iter_parseable(stream: &[u8]) {
    let mut iter = records::BiffRecordIter::from_offset(stream, 0).expect("offset 0 in-bounds");
    while let Some(next) = iter.next() {
        assert!(next.is_ok(), "physical iterator errored: {next:?}");
    }
}

fn record_spec_strategy() -> impl Strategy<Value = RecordSpec> {
    // Avoid BOF/EOF/FILEPASS so the generator can control substream structure.
    let id = any::<u16>().prop_filter("exclude special record ids", |id| {
        !matches!(
            *id,
            records::RECORD_BOF_BIFF8
                | records::RECORD_BOF_BIFF5
                | records::RECORD_EOF
                | records::RECORD_FILEPASS
        )
    });

    let payload = proptest::collection::vec(any::<u8>(), 0..=MAX_RECORD_PAYLOAD_LEN);
    (id, payload).prop_map(|(record_id, payload)| RecordSpec { record_id, payload })
}

fn workbook_stream_strategy(include_filepass: bool) -> impl Strategy<Value = GeneratedStream> {
    let globals = proptest::collection::vec(record_spec_strategy(), 0..=MAX_GLOBAL_RECORDS);

    globals.prop_flat_map(move |globals_records| {
        let idx_range = 0usize..=globals_records.len();
        (
            Just(globals_records),
            idx_range,
            proptest::collection::vec(any::<u8>(), 0..=MAX_FILEPASS_PAYLOAD_LEN),
            any::<bool>(),
            proptest::collection::vec(record_spec_strategy(), 0..=MAX_SHEET_RECORDS),
        )
            .prop_map(
                move |(
                    globals_records,
                    filepass_idx,
                    filepass_payload,
                    include_sheet,
                    sheet_records,
                )| {
                    let mut bytes = Vec::new();
                    let mut payload_ranges_after_filepass = Vec::new();
                    let mut after_filepass = false;
                    let mut filepass_present = false;

                    // Workbook globals BOF.
                    let bof_globals_payload = [0x00, 0x06, 0x05, 0x00]; // BIFF8 + workbook globals
                    push_record(&mut bytes, records::RECORD_BOF_BIFF8, &bof_globals_payload);

                    for (idx, rec) in globals_records.iter().enumerate() {
                        if include_filepass && idx == filepass_idx {
                            push_record(&mut bytes, records::RECORD_FILEPASS, &filepass_payload);
                            after_filepass = true;
                            filepass_present = true;
                        }

                        let record_offset = push_record(&mut bytes, rec.record_id, &rec.payload);
                        if after_filepass {
                            payload_ranges_after_filepass
                                .push((record_offset + 4, rec.payload.len()));
                        }
                    }

                    if include_filepass && filepass_idx == globals_records.len() {
                        push_record(&mut bytes, records::RECORD_FILEPASS, &filepass_payload);
                        after_filepass = true;
                        filepass_present = true;
                    }

                    // Workbook globals EOF.
                    let eof_offset = push_record(&mut bytes, records::RECORD_EOF, &[]);
                    if after_filepass {
                        payload_ranges_after_filepass.push((eof_offset + 4, 0));
                    }

                    // Optional second substream (sheet).
                    if include_sheet {
                        let bof_sheet_payload = [0x00, 0x06, 0x10, 0x00]; // BIFF8 + worksheet
                        let bof_offset =
                            push_record(&mut bytes, records::RECORD_BOF_BIFF8, &bof_sheet_payload);
                        if after_filepass {
                            payload_ranges_after_filepass
                                .push((bof_offset + 4, bof_sheet_payload.len()));
                        }
                        for rec in sheet_records.iter() {
                            let record_offset =
                                push_record(&mut bytes, rec.record_id, &rec.payload);
                            if after_filepass {
                                payload_ranges_after_filepass
                                    .push((record_offset + 4, rec.payload.len()));
                            }
                        }
                        let eof_offset = push_record(&mut bytes, records::RECORD_EOF, &[]);
                        if after_filepass {
                            payload_ranges_after_filepass.push((eof_offset + 4, 0));
                        }
                    }

                    GeneratedStream {
                        bytes,
                        filepass_present,
                        payload_ranges_after_filepass,
                    }
                },
            )
    })
}

fn any_workbook_stream() -> impl Strategy<Value = GeneratedStream> {
    any::<bool>().prop_flat_map(|include_filepass| workbook_stream_strategy(include_filepass))
}

fn workbook_stream_with_filepass() -> impl Strategy<Value = GeneratedStream> {
    workbook_stream_strategy(true).prop_filter("must contain FILEPASS", |s| s.filepass_present)
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 128,
        rng_seed: proptest::test_runner::RngSeed::Fixed(0xC0FFEE),
        failure_persistence: None,
        .. ProptestConfig::default()
    })]

    #[test]
    fn decryptor_roundtrips_xor_and_mask_preserves_record_boundaries(case in any_workbook_stream()) {
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                // Basic sanity: our generator should always emit a structurally valid BIFF stream.
                assert_physical_iter_parseable(&case.bytes);

                // (1) Decryptor record-walking logic is panic-free and does not corrupt headers.
                let plaintext = case.bytes.clone();
                let mut ciphertext = plaintext.clone();

                if case.filepass_present {
                    // Apply the spec-defined XOR "encryption" to produce ciphertext, then decrypt it
                    // using the record-walker. If the walker accidentally decrypts record headers or
                    // miscounts block offsets, the roundtrip equality check will fail.
                    apply_xor_spec(&mut ciphertext, &case.payload_ranges_after_filepass);

                    decrypt_workbook_stream_with_cipher(&mut ciphertext, "pw", xor_cipher)
                        .expect("decrypt should succeed for structurally valid streams with FILEPASS");
                    assert_eq!(ciphertext, plaintext, "decryptor did not restore the original stream");

                    // Decryption must not break record boundaries (headers are plaintext).
                    assert_physical_iter_parseable(&ciphertext);
                } else {
                    let err = decrypt_workbook_stream_with_cipher(&mut ciphertext, "pw", xor_cipher)
                        .expect_err("expected NoFilePass for streams without FILEPASS");
                    assert_eq!(err, DecryptError::NoFilePass);
                }

                // (2) FILEPASS masking preserves stream length and BIFF record boundaries.
                let mut masked = plaintext.clone();
                let orig_len = masked.len();
                let orig_meta = collect_physical_meta(&masked);

                let masked_count = records::mask_workbook_globals_filepass_record_id_in_place(&mut masked);
                assert_eq!(masked_count > 0, case.filepass_present);
                assert_eq!(masked.len(), orig_len, "masking must not change stream length");
                assert_physical_iter_parseable(&masked);

                // Physical record boundaries (offset + payload length) must be unchanged.
                let new_meta = collect_physical_meta(&masked);
                assert_eq!(new_meta.len(), orig_meta.len());
                for (old, new) in orig_meta.iter().zip(new_meta.iter()) {
                    assert_eq!(old.0, new.0, "record offset changed after masking");
                    assert_eq!(old.2, new.2, "record length changed after masking");
                }

                // After masking, the globals scan must not report FILEPASS even if it was present.
                assert!(
                    !records::workbook_globals_has_filepass_record(&masked),
                    "masked stream should not be detected as encrypted"
                );
            }))
            .is_ok(),
            "panic in decryptor/masking logic",
        );
    }
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 64,
        rng_seed: proptest::test_runner::RngSeed::Fixed(0xBADF00D),
        failure_persistence: None,
        .. ProptestConfig::default()
    })]

    #[test]
    fn decryptor_errors_on_truncated_or_corrupt_stream(case in workbook_stream_with_filepass()) {
        prop_assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                // Start from a ciphertext stream so the decryptor actually exercises the post-FILEPASS
                // record walker.
                let mut ciphertext = case.bytes.clone();
                apply_xor_spec(&mut ciphertext, &case.payload_ranges_after_filepass);

                // Malformed variant 1: truncate the stream (forces a truncated header or payload).
                if ciphertext.len() > 0 {
                    let mut truncated = ciphertext.clone();
                    truncated.pop();
                    let err = decrypt_workbook_stream_with_cipher(&mut truncated, "pw", xor_cipher)
                        .expect_err("expected decryptor to error on truncated stream");
                    let _ = err; // Any error is acceptable; we only require panic-free behavior.
                }

                // Malformed variant 2: corrupt a length field so a record extends past end-of-stream.
                let mut corrupt_len = ciphertext.clone();
                let eof_offset = {
                    let mut iter = records::BiffRecordIter::from_offset(&corrupt_len, 0).expect("offset 0 in-bounds");
                    let mut last: Option<records::BiffRecord<'_>> = None;
                    while let Some(next) = iter.next() {
                        match next {
                            Ok(record) => last = Some(record),
                            Err(_) => break,
                        }
                    }
                    last.expect("valid stream must have at least one record").offset
                };

                // Overwrite the payload length to a large value (u16::MAX) without adding bytes.
                // This must trigger a bounds error in the physical iterator.
                if eof_offset + 4 <= corrupt_len.len() {
                    corrupt_len[eof_offset + 2..eof_offset + 4].copy_from_slice(&u16::MAX.to_le_bytes());
                }

                let err = decrypt_workbook_stream_with_cipher(&mut corrupt_len, "pw", xor_cipher)
                    .expect_err("expected decryptor to error on corrupt length");
                let _ = err;
            }))
            .is_ok(),
            "panic in decryptor on malformed input",
        );
    }
}
