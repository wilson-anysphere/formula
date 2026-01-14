//! Utility for generating Office-encrypted OOXML fixtures (OLE `EncryptedPackage` wrapper).
//!
//! This is **not** used by production code; it's a small helper for regenerating test fixtures
//! without relying on Excel/Office.

use formula_office_crypto::{encrypt_package_to_ole, EncryptOptions};

fn main() {
    let mut args = std::env::args().skip(1).collect::<Vec<_>>();
    if args.len() != 3 {
        eprintln!("Usage: encrypt_ooxml_agile <input.xlsm> <output.xlsm> <password>");
        std::process::exit(2);
    }

    let password = args.pop().expect("password");
    let output = args.pop().expect("output");
    let input = args.pop().expect("input");

    let plaintext = std::fs::read(&input).expect("read input");
    let encrypted =
        encrypt_package_to_ole(&plaintext, &password, EncryptOptions::default()).expect("encrypt");
    std::fs::write(&output, encrypted).expect("write output");
}

