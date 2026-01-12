use std::ffi::OsString;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use encoding_rs::UTF_16LE;
use formula_vba::{decompress_container, OleFile};

fn main() -> ExitCode {
    match run() {
        Ok(()) => ExitCode::SUCCESS,
        Err(err) => {
            eprintln!("{err}");
            ExitCode::FAILURE
        }
    }
}

fn run() -> Result<(), String> {
    let mut args = std::env::args_os();
    let program = args
        .next()
        .unwrap_or_else(|| OsString::from("dump_dir_records"));

    let Some(input) = args.next() else {
        return Err(usage(&program));
    };
    if args.next().is_some() {
        return Err(usage(&program));
    }

    let input_path = PathBuf::from(input);

    let (vba_project_bin, source) = load_vba_project_bin(&input_path)?;

    println!("vbaProject.bin source: {source}");
    println!("vbaProject.bin size: {} bytes", vba_project_bin.len());

    let mut ole =
        OleFile::open(&vba_project_bin).map_err(|e| format!("failed to parse OLE: {e}"))?;

    let dir_compressed = ole
        .read_stream_opt("VBA/dir")
        .map_err(|e| format!("failed to read VBA/dir: {e}"))?
        .ok_or("missing required stream VBA/dir".to_owned())?;

    println!("VBA/dir compressed: {} bytes", dir_compressed.len());

    let dir_decompressed = decompress_container(&dir_compressed)
        .map_err(|e| format!("failed to decompress VBA/dir container: {e}"))?;

    println!("VBA/dir decompressed: {} bytes", dir_decompressed.len());
    println!();
    println!("-- VBA/dir records (decompressed) --");
    dump_dir_records(&dir_decompressed);

    dump_project_normalized_data_v3_dir_records(&vba_project_bin);
    dump_project_normalized_data_v3(&vba_project_bin);

    Ok(())
}

fn usage(program: &OsString) -> String {
    format!(
        "usage: {} <vbaProject.bin|workbook.xlsm|workbook.xlsx|workbook.xlsb>",
        program.to_string_lossy()
    )
}

fn load_vba_project_bin(path: &Path) -> Result<(Vec<u8>, String), String> {
    match try_extract_vba_project_bin_from_zip(path) {
        Ok(Some(bytes)) => Ok((
            bytes,
            format!("{} (zip entry xl/vbaProject.bin)", path.display()),
        )),
        Ok(None) => {
            // Not a zip workbook; treat as a raw vbaProject.bin OLE file.
            let bytes = std::fs::read(path)
                .map_err(|e| format!("failed to read {}: {e}", path.display()))?;
            Ok((bytes, path.display().to_string()))
        }
        Err(err) => Err(err),
    }
}

fn try_extract_vba_project_bin_from_zip(path: &Path) -> Result<Option<Vec<u8>>, String> {
    let file = match std::fs::File::open(path) {
        Ok(f) => f,
        Err(e) => return Err(format!("failed to open {}: {e}", path.display())),
    };

    let mut archive = match zip::ZipArchive::new(file) {
        Ok(a) => a,
        Err(_) => return Ok(None),
    };

    let mut entry = match archive.by_name("xl/vbaProject.bin") {
        Ok(f) => f,
        Err(zip::result::ZipError::FileNotFound) => {
            return Err(format!(
                "{} is a zip, but does not contain xl/vbaProject.bin",
                path.display()
            ));
        }
        Err(e) => return Err(format!("failed to read zip {}: {e}", path.display())),
    };

    let mut buf = Vec::new();
    entry
        .read_to_end(&mut buf)
        .map_err(|e| format!("failed to read xl/vbaProject.bin from {}: {e}", path.display()))?;
    Ok(Some(buf))
}

fn dump_dir_records(decompressed: &[u8]) {
    let mut offset = 0usize;
    let mut idx = 0usize;

    while offset < decompressed.len() {
        let record_offset = offset;

        if offset + 6 > decompressed.len() {
            println!(
                "[{:03}] offset=0x{record_offset:08x} <truncated record header: need 6 bytes, have {}>",
                idx + 1,
                decompressed.len() - offset
            );
            break;
        }

        let id = u16::from_le_bytes([decompressed[offset], decompressed[offset + 1]]);
        let len = u32::from_le_bytes([
            decompressed[offset + 2],
            decompressed[offset + 3],
            decompressed[offset + 4],
            decompressed[offset + 5],
        ]) as usize;
        offset += 6;

        if offset + len > decompressed.len() {
            println!(
                "[{:03}] offset=0x{record_offset:08x} id={id:#06x} len={len} <bad record length: need {} bytes, have {}>",
                idx + 1,
                len,
                decompressed.len().saturating_sub(offset)
            );
            break;
        }
        let data = &decompressed[offset..offset + len];
        offset += len;

        idx += 1;
        let name = record_name(id).unwrap_or("<unknown>");
        println!(
            "[{idx:03}] offset=0x{record_offset:08x} id={id:#06x} len={len:>6} {name}"
        );

        if len <= 64 {
            println!("      hex: {}", bytes_to_hex_spaced(data));
            if !data.is_empty() {
                println!("      ascii: {}", bytes_to_ascii_preview(data));
            }
            if let Some(v) = record_numeric_preview(id, data) {
                println!("      value: {v}");
            }

            // MS-OVBA v3 "Unicode" record variants often store:
            //   u32 length prefix (code units or bytes) || UTF-16LE bytes
            // The generic UTF-16LE heuristic will mis-decode the length prefix as a character, so
            // handle these record ids specially.
            if is_len_prefixed_unicode_record_id(id) {
                // Prefer decoding just the payload (skipping the internal u32 length prefix) for the
                // v3 Unicode record variants, but fall back to decoding the raw bytes if the record
                // does not match the expected shape. This keeps the output useful even when we are
                // pointed at non-canonical or partially malformed inputs.
                let (bytes_to_decode, label) = match unicode_record_payload_len_prefixed(data) {
                    Some(payload) => (payload, "utf16le(payload)"),
                    None => (data, "utf16le(raw)"),
                };
                if looks_like_utf16le(bytes_to_decode) {
                    let (cow, had_errors) = UTF_16LE.decode_without_bom_handling(bytes_to_decode);
                    let mut s = cow.into_owned();
                    s.retain(|c| c != '\u{0000}');
                    let escaped = escape_str(&s);
                    if had_errors {
                        println!("      {label}: {escaped} <decode errors>");
                    } else {
                        println!("      {label}: {escaped}");
                    }
                }
            } else if looks_like_utf16le(data) {
                let (cow, had_errors) = UTF_16LE.decode_without_bom_handling(data);
                let mut s = cow.into_owned();
                // This is a debugging aid: strip NULs to keep output readable.
                s.retain(|c| c != '\u{0000}');
                let escaped = escape_str(&s);
                if had_errors {
                    println!("      utf16le: {escaped} <decode errors>");
                } else {
                    println!("      utf16le: {escaped}");
                }
            }
        }
    }

    if offset == decompressed.len() {
        println!();
        println!("records: {idx}");
    } else {
        println!();
        println!("records: {idx} (stopped early at offset=0x{offset:08x})");
    }
}

fn record_name(id: u16) -> Option<&'static str> {
    // Names from MS-OVBA 2.3.4 "dir Stream".
    Some(match id {
        // ---- Project information records ----
        0x0001 => "PROJECTSYSKIND",
        0x0002 => "PROJECTLCID",
        0x0003 => "PROJECTCODEPAGE",
        0x0004 => "PROJECTNAME",
        0x0005 => "PROJECTDOCSTRING",
        0x0006 => "PROJECTHELPFILEPATH",
        0x0007 => "PROJECTHELPCONTEXT",
        0x0008 => "PROJECTLIBFLAGS",
        0x0009 => "PROJECTVERSION",
        0x000C => "PROJECTCONSTANTS",
        0x0014 => "PROJECTLCIDINVOKE",
        // Unicode variants of selected strings.
        //
        // MS-OVBA v3 `ProjectNormalizedData` uses the 0x0040..0x0043 Unicode record IDs with an
        // *internal* u32 length prefix (see docs/vba-digital-signatures.md).
        0x0040 => "PROJECTNAMEUNICODE",
        0x0041 => "PROJECTDOCSTRINGUNICODE",
        0x0042 => "PROJECTHELPFILEPATHUNICODE",
        0x0043 => "PROJECTCONSTANTSUNICODE",
        // Legacy/alternate record id seen in some older fixtures/implementations.
        0x003C => "PROJECTCONSTANTSUNICODE (legacy id 0x003C)",
        // Present in many real-world files, but skipped by the MS-OVBA V3ContentNormalizedData
        // pseudocode.
        0x004A => "PROJECTCOMPATVERSION / MODULEHELPFILEPATHUNICODE (id collision)",

        // ---- Reference records (used by ContentNormalizedData / ProjectNormalizedData) ----
        0x000D => "REFERENCEREGISTERED",
        0x000E => "REFERENCEPROJECT",
        0x0016 => "REFERENCENAME",
        0x002F => "REFERENCECONTROL",
        0x0030 => "REFERENCEEXTENDED",
        0x0033 => "REFERENCEORIGINAL",

        // ---- Module records ----
        // ProjectModules / terminators (spec-accurate `VBA/dir` layouts use these).
        0x000F => "PROJECTMODULES (ModuleCount)",
        0x0013 => "PROJECTCOOKIE",
        0x0010 => "PROJECTTERMINATOR (dir stream end)",

        0x0019 => "MODULENAME",
        0x0047 => "MODULENAMEUNICODE",
        0x001B => "MODULEDOCSTRING",
        0x001A => "MODULESTREAMNAME",
        // MS-OVBA v3 Unicode variant id for MODULESTREAMNAME (with internal u32 length prefix).
        0x0048 => "MODULESTREAMNAMEUNICODE",
        // MS-OVBA v3 Unicode variant id for MODULEDOCSTRING (with internal u32 length prefix).
        0x0049 => "MODULEDOCSTRINGUNICODE",
        // Legacy/alternate ids seen in some fixtures/implementations.
        0x0032 => "MODULESTREAMNAMEUNICODE (legacy id 0x0032)",
        0x001C => "MODULEDOCSTRING (legacy id 0x001C)",
        0x001D => "MODULEHELPFILEPATH",
        0x001E => "MODULEHELPCONTEXT",
        0x0021 => "MODULETYPE (procedural TypeRecord.Id=0x0021)",
        0x0022 => "MODULETYPE (non-procedural TypeRecord.Id=0x0022)",
        0x0025 => "MODULEREADONLY",
        0x0028 => "MODULEPRIVATE",
        0x002B => "MODULETERMINATOR",
        0x002C => "MODULECOOKIE",
        0x0031 => "MODULETEXTOFFSET",

        _ => return None,
    })
}

fn bytes_to_hex_spaced(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len() * 3);
    for (i, b) in bytes.iter().enumerate() {
        use std::fmt::Write;
        if i != 0 {
            out.push(' ');
        }
        write!(&mut out, "{:02x}", b).expect("writing to String cannot fail");
    }
    out
}

fn bytes_to_ascii_preview(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len());
    for &b in bytes {
        match b {
            b'\r' => out.push_str("\\r"),
            b'\n' => out.push_str("\\n"),
            b'\t' => out.push_str("\\t"),
            0x20..=0x7E => out.push(b as char),
            0x00 => out.push_str("\\0"),
            _ => out.push('.'),
        }
    }
    out
}

fn record_numeric_preview(id: u16, data: &[u8]) -> Option<String> {
    match id {
        // PROJECTSYSKIND / PROJECTLCID / PROJECTLCIDINVOKE / PROJECTHELPCONTEXT / PROJECTLIBFLAGS
        0x0001 | 0x0002 | 0x0014 | 0x0007 | 0x0008 if data.len() >= 4 => {
            let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
            Some(format!("u32_le={n} (0x{n:08x})"))
        }
        // PROJECTCODEPAGE
        0x0003 if data.len() >= 2 => {
            let n = u16::from_le_bytes([data[0], data[1]]);
            Some(format!("u16_le={n} (0x{n:04x})"))
        }
        // PROJECTVERSION is commonly `u16 major || u16 minor`, but some producers embed a longer
        // structure. Print a best-effort interpretation.
        0x0009 if data.len() == 4 => {
            let major = u16::from_le_bytes([data[0], data[1]]);
            let minor = u16::from_le_bytes([data[2], data[3]]);
            Some(format!("version={major}.{minor} (u16/u16)"))
        }
        0x0009 if data.len() >= 10 => {
            let reserved = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
            let major = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
            let minor = u16::from_le_bytes([data[8], data[9]]);
            Some(format!(
                "reserved=0x{reserved:08x} major={major} minor={minor} (u32/u32/u16)"
            ))
        }
        // MODULETYPE
        0x0021 | 0x0022 if data.len() >= 2 => {
            let n = u16::from_le_bytes([data[0], data[1]]);
            Some(format!("u16_le={n} (0x{n:04x})"))
        }
        // MODULETEXTOFFSET
        0x0031 if data.len() >= 4 => {
            let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
            Some(format!("u32_le={n} (0x{n:08x})"))
        }
        _ => None,
    }
}

fn looks_like_utf16le(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || bytes.len() % 2 != 0 {
        return false;
    }
    // If a substantial portion of the high bytes are NUL, it's probably
    // UTF-16LE for ASCII-range characters.
    let total = bytes.len() / 2;
    let nul_high = bytes.iter().skip(1).step_by(2).filter(|&&b| b == 0).count();
    // Use a ceiling-half threshold: for very short inputs (e.g. 2 bytes), `total / 2` is 0 and
    // would incorrectly classify any 2-byte MBCS string as UTF-16LE.
    nul_high * 2 >= total
}

fn escape_str(s: &str) -> String {
    s.chars().flat_map(|c| c.escape_default()).collect()
}

fn is_len_prefixed_unicode_record_id(id: u16) -> bool {
    matches!(
        id,
        // Project Unicode string variants (MS-OVBA v3).
        0x0040 | 0x0041 | 0x0042 | 0x0043
            // Module Unicode string variants (MS-OVBA v3).
            | 0x0047 | 0x0048 | 0x0049
            // 0x004A is an ID collision: PROJECTCOMPATVERSION vs MODULEHELPFILEPATHUNICODE.
            | 0x004A
    )
}

fn unicode_record_payload_len_prefixed(data: &[u8]) -> Option<&[u8]> {
    if data.len() < 4 {
        return None;
    }
    let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
    let remaining = data.len() - 4;

    let bytes_by_units = n.checked_mul(2);

    // Prefer interpretations that exactly match the record length.
    if let Some(bytes) = bytes_by_units {
        if bytes == remaining {
            return Some(&data[4..4 + bytes]);
        }
    }
    if n == remaining {
        return Some(&data[4..4 + n]);
    }

    // Otherwise, prefer UTF-16 code unit count when it fits; fall back to byte count.
    if let Some(bytes) = bytes_by_units {
        if bytes <= remaining {
            return Some(&data[4..4 + bytes]);
        }
    }
    if n <= remaining {
        return Some(&data[4..4 + n]);
    }

    None
}

fn dump_project_normalized_data_v3_dir_records(vba_project_bin: &[u8]) {
    const PREFIX_LEN: usize = 64;
    println!();
    println!("-- ProjectNormalizedDataV3 (dir records) --");

    match formula_vba::project_normalized_data_v3_dir_records(vba_project_bin) {
        Ok(data) => {
            let n = PREFIX_LEN.min(data.len());
            println!("len: {} bytes", data.len());
            println!("first {n} bytes: {}", bytes_to_hex_spaced(&data[..n]));
        }
        Err(err) => {
            // Keep going: this is a developer tool and should be resilient to partially malformed
            // inputs / in-progress implementations.
            println!("error: {err}");
        }
    }
}

fn dump_project_normalized_data_v3(vba_project_bin: &[u8]) {
    const PREFIX_LEN: usize = 64;
    println!();
    println!(
        "-- ProjectNormalizedDataV3 (filtered PROJECT props || V3ContentNormalizedData || FormsNormalizedData) --"
    );

    match formula_vba::project_normalized_data_v3(vba_project_bin) {
        Ok(data) => {
            let n = PREFIX_LEN.min(data.len());
            println!("len: {} bytes", data.len());
            println!("first {n} bytes: {}", bytes_to_hex_spaced(&data[..n]));
        }
        Err(err) => {
            // Keep going: this is a developer tool and should be resilient to partially malformed
            // inputs / in-progress implementations.
            println!("error: {err}");
        }
    }
}
