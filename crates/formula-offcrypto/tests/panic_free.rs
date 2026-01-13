use formula_offcrypto::{parse_encrypted_package_header, parse_encryption_info};

fn next_u64(state: &mut u64) -> u64 {
    // Deterministic LCG (same parameters as PCG32 without the output permutation).
    *state = state
        .wrapping_mul(6364136223846793005)
        .wrapping_add(1442695040888963407);
    *state
}

#[test]
fn parsers_are_panic_free_on_pseudorandom_inputs() {
    let mut state = 0x0123_4567_89ab_cdef;

    for _ in 0..1024 {
        let len = (next_u64(&mut state) as usize) % 8192;
        let mut buf = vec![0u8; len];
        for b in &mut buf {
            *b = (next_u64(&mut state) >> 56) as u8;
        }

        assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = parse_encryption_info(&buf);
            }))
            .is_ok(),
            "parse_encryption_info panicked on len={len}"
        );

        assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let _ = parse_encrypted_package_header(&buf);
            }))
            .is_ok(),
            "parse_encrypted_package_header panicked on len={len}"
        );
    }
}

