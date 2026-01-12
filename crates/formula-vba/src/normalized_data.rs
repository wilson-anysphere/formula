use std::io::{Cursor, Read};

use crate::OleError;

/// Compute the MS-OVBA §2.4.2.2 `FormsNormalizedData` byte sequence for a `vbaProject.bin`.
///
/// This is used as input to the "Agile" VBA project signature binding algorithm.
///
/// Behavior (matching MS-OVBA intent, and kept deterministic for hashing):
/// - Enumerate all **streams** contained in root-level "designer" storages (i.e. storages that are
///   not the `VBA` storage and not a `\x05DigitalSignature*` storage).
/// - Recursively include streams from nested storages.
/// - Streams are concatenated in **lexicographic order by full OLE path** (e.g. `Form/Child/X`
///   comes before `Form/Y`). This avoids relying on OLE directory entry ordering details.
/// - Each stream's bytes are padded with `0x00` bytes up to a multiple of **1023 bytes**
///   (MS-OVBA splits stream data into 1023-byte blocks and zero-pads the final block).
pub fn forms_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, OleError> {
    // `cfb` works over any `Read + Seek`; use a borrowed cursor to avoid copying the input bytes.
    let cursor = Cursor::new(vba_project_bin);
    let mut file = cfb::CompoundFile::open(cursor)?;

    let mut designer_stream_paths = Vec::new();
    for entry in file.walk() {
        if !entry.is_stream() {
            continue;
        }

        let path = entry.path().to_string_lossy().to_string();
        let Some((root_component, _rest)) = path.split_once('/') else {
            // Root-level streams (e.g. `PROJECT`) are not part of `FormsNormalizedData`.
            continue;
        };

        if root_component == "VBA" {
            continue;
        }
        if is_signature_component(root_component) {
            continue;
        }

        designer_stream_paths.push(path);
    }

    // Deterministic order for hashing.
    designer_stream_paths.sort();

    let mut out = Vec::new();
    for path in designer_stream_paths {
        let mut s = file.open_stream(&path)?;
        let mut buf = Vec::new();
        s.read_to_end(&mut buf)?;

        out.extend_from_slice(&buf);

        let rem = buf.len() % 1023;
        if rem != 0 {
            out.extend(std::iter::repeat(0u8).take(1023 - rem));
        }
    }

    Ok(out)
}

fn is_signature_component(component: &str) -> bool {
    // Excel/VBA signature storages/streams are control-character prefixed in OLE; normalize by
    // stripping leading C0 control chars.
    let trimmed = component.trim_start_matches(|c: char| c <= '\u{001F}');
    matches!(
        trimmed,
        "DigitalSignature" | "DigitalSignatureEx" | "DigitalSignatureExt"
    )
}

