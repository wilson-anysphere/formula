#![allow(unexpected_cfgs)]

use proptest::prelude::*;

use super::*;

// Keep CI runtime bounded. Heavier fuzzing can be enabled by building with
// `RUSTFLAGS=\"--cfg fuzzing\"` (or an equivalent `cfg(fuzzing)` setup).
#[cfg(fuzzing)]
const CASES: u32 = 1024;
#[cfg(not(fuzzing))]
const CASES: u32 = 64;

#[cfg(fuzzing)]
const MAX_INPUT_LEN: usize = 256 * 1024;
#[cfg(not(fuzzing))]
const MAX_INPUT_LEN: usize = 32 * 1024;

proptest! {
    #![proptest_config(ProptestConfig {
        cases: CASES,
        .. ProptestConfig::default()
    })]

    #[test]
    fn parse_encryption_info_agile_is_panic_free_and_rejects_malformed_xml(tail in proptest::collection::vec(any::<u8>(), 0..=MAX_INPUT_LEN)) {
        // Ensure this is not accidentally a valid XML document (which could cause a rare `Ok` and
        // make the property test flaky). Inject a byte sequence that is never valid UTF-8.
        let mut bytes = Vec::with_capacity(8 + 2 + tail.len());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
        bytes.push(b'<');
        bytes.push(0xFF);
        bytes.extend_from_slice(&tail);

        let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| parse_encryption_info(&bytes)));
        prop_assert!(res.is_ok(), "parse_encryption_info panicked");

        let parsed = res.unwrap();
        prop_assert!(parsed.is_err(), "expected malformed agile XML to be rejected");
    }
}
