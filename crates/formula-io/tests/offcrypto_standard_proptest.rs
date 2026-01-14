#![cfg(all(not(target_arch = "wasm32"), feature = "offcrypto"))]

use formula_io::offcrypto::standard::{parse_encryption_info_standard, verify_password_standard};
use proptest::prelude::*;

proptest! {
    // Keep the fuzz surface modest so the suite is fast and stable in CI.
    #![proptest_config(ProptestConfig {
        cases: 256,
        // Keep fuzz-style tests deterministic in CI so failures are reproducible.
        rng_seed: proptest::test_runner::RngSeed::Fixed(0),
        max_shrink_iters: 0,
        failure_persistence: None,
        .. ProptestConfig::default()
    })]

    #[test]
    fn standard_encryptioninfo_parse_and_verify_never_panics(
        bytes in proptest::collection::vec(any::<u8>(), 0..=4096)
    ) {
        let parsed = std::panic::catch_unwind(|| parse_encryption_info_standard(&bytes));
        prop_assert!(parsed.is_ok(), "parse_encryption_info_standard panicked");

        match parsed.unwrap() {
            Ok(info) => {
                let verified = std::panic::catch_unwind(|| verify_password_standard(&info, ""));
                prop_assert!(verified.is_ok(), "verify_password_standard panicked");
                prop_assert!(verified.unwrap().is_ok(), "verify_password_standard returned an error");
            }
            Err(_err) => {
                // Structured error is fine; the invariant we care about is "no panics".
            }
        }
    }
}
