use std::ffi::OsString;
use std::path::PathBuf;
use std::process::ExitCode;

use encoding_rs::UTF_16LE;
use formula_vba::{decompress_container, OleFile};

#[path = "shared/dir_record_names.rs"]
mod dir_record_names;
#[path = "shared/vba_project_bin.rs"]
mod vba_project_bin;
#[path = "shared/broken_pipe.rs"]
mod broken_pipe;

fn main() -> ExitCode {
    broken_pipe::install();
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

    let (input_path, password) = parse_args(&program, args)?;

    let (vba_project_bin, source) =
        vba_project_bin::load_vba_project_bin(&input_path, password.as_deref())?;

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
        "usage: {} --input <vbaProject.bin|workbook.xlsm|workbook.xlsx|workbook.xlsb> [--password <password>|--password-file <path>]",
        program.to_string_lossy()
    )
}

fn parse_args(
    program: &OsString,
    args: impl Iterator<Item = OsString>,
) -> Result<(PathBuf, Option<String>), String> {
    let mut args = args.peekable();
    let mut input: Option<PathBuf> = None;
    let mut password: Option<String> = None;
    let mut password_file: Option<PathBuf> = None;

    while let Some(arg) = args.next() {
        let arg_str = arg.to_string_lossy();

        if arg_str == "--help" || arg_str == "-h" {
            return Err(usage(program));
        }

        if arg_str == "--input" {
            let value = args.next().ok_or_else(|| usage(program))?;
            input = Some(PathBuf::from(value));
            continue;
        }
        if let Some(rest) = arg_str.strip_prefix("--input=") {
            input = Some(PathBuf::from(rest));
            continue;
        }

        if arg_str == "--password" {
            let value = args.next().ok_or_else(|| usage(program))?;
            password = Some(value.to_string_lossy().into_owned());
            continue;
        }
        if let Some(rest) = arg_str.strip_prefix("--password=") {
            password = Some(rest.to_string());
            continue;
        }

        if arg_str == "--password-file" {
            let value = args.next().ok_or_else(|| usage(program))?;
            password_file = Some(PathBuf::from(value));
            continue;
        }
        if let Some(rest) = arg_str.strip_prefix("--password-file=") {
            password_file = Some(PathBuf::from(rest));
            continue;
        }

        // Backwards-compatible positional input path.
        if input.is_none() {
            input = Some(PathBuf::from(arg));
            continue;
        }

        return Err(usage(program));
    }

    if password.is_some() && password_file.is_some() {
        return Err(format!(
            "use only one of --password or --password-file\n{}",
            usage(program)
        ));
    }

    let password = match password_file {
        Some(path) => {
            let text = std::fs::read_to_string(&path)
                .map_err(|e| format!("failed to read password file {}: {e}", path.display()))?;
            Some(text.trim_end_matches(['\r', '\n']).to_string())
        }
        None => password,
    };

    let Some(input) = input else {
        return Err(usage(program));
    };
    Ok((input, password))
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
        let len_field = u32::from_le_bytes([
            decompressed[offset + 2],
            decompressed[offset + 3],
            decompressed[offset + 4],
            decompressed[offset + 5],
        ]) as usize;

        // Most `VBA/dir` records are encoded as `Id(u16) || Size(u32) || Data(Size)`.
        //
        // However, real-world projects can store PROJECTVERSION (0x0009) in the fixed-length form:
        //   Id(u16) || Reserved(u32) || VersionMajor(u32) || VersionMinor(u16)
        // (12 bytes total).
        //
        // Many of our fixtures and some producers instead encode it as TLV (`Id || Size || Data`).
        // Disambiguate by checking which interpretation yields a plausible next record boundary.
        let (data, record_len, next_offset) = if id == 0x0009 {
            let tlv_end = record_offset.saturating_add(6).saturating_add(len_field);
            let fixed_end = record_offset.saturating_add(12);

            let tlv_next_ok = looks_like_projectversion_following_record(decompressed, tlv_end);
            let fixed_next_ok = looks_like_projectversion_following_record(decompressed, fixed_end);

            if len_field == 0 && fixed_end > decompressed.len() {
                println!(
                    "[{:03}] offset=0x{record_offset:08x} id={id:#06x} <truncated PROJECTVERSION: need 12 bytes, have {}>",
                    idx + 1,
                    decompressed.len() - record_offset
                );
                break;
            }

            if fixed_end <= decompressed.len() && fixed_next_ok && (!tlv_next_ok || len_field == 0)
            {
                let data = &decompressed[record_offset + 2..fixed_end];
                (data, data.len(), fixed_end)
            } else {
                let data_start = record_offset + 6;
                let data_end = data_start.saturating_add(len_field);
                if data_end > decompressed.len() {
                    println!(
                        "[{:03}] offset=0x{record_offset:08x} id={id:#06x} len={len_field} <bad record length: need {} bytes, have {}>",
                        idx + 1,
                        len_field,
                        decompressed.len().saturating_sub(data_start)
                    );
                    break;
                }
                (&decompressed[data_start..data_end], len_field, data_end)
            }
        } else {
            let data_start = record_offset + 6;
            let data_end = data_start.saturating_add(len_field);
            if data_end > decompressed.len() {
                println!(
                    "[{:03}] offset=0x{record_offset:08x} id={id:#06x} len={len_field} <bad record length: need {} bytes, have {}>",
                    idx + 1,
                    len_field,
                    decompressed.len().saturating_sub(data_start)
                );
                break;
            }
            (&decompressed[data_start..data_end], len_field, data_end)
        };

        offset = next_offset;

        idx += 1;
        let name = dir_record_names::record_name(id).unwrap_or("<unknown>");
        println!("[{idx:03}] offset=0x{record_offset:08x} id={id:#06x} len={record_len:>6} {name}");

        if record_len <= 64 {
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
                if !bytes_to_decode.is_empty() && bytes_to_decode.len() % 2 == 0 {
                    // For known Unicode record IDs, it's more useful to decode unconditionally than
                    // to rely on heuristics based on NUL high bytes (which only works for
                    // ASCII-range strings).
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

fn looks_like_projectversion_following_record(bytes: &[u8], offset: usize) -> bool {
    if offset == bytes.len() {
        return true;
    }
    if offset + 6 > bytes.len() {
        return false;
    }
    let id = u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
    if !matches!(
        id,
        0x000C
            | 0x003C
            | 0x004A
            | 0x000D
            | 0x000E
            | 0x0016
            | 0x002F
            | 0x0030
            | 0x0033
            | 0x000F
            | 0x0013
            | 0x0010
            | 0x0019
            | 0x0047
            | 0x001A
            | 0x0032
    ) {
        return false;
    }
    let len = u32::from_le_bytes([
        bytes[offset + 2],
        bytes[offset + 3],
        bytes[offset + 4],
        bytes[offset + 5],
    ]) as usize;
    offset + 6 + len <= bytes.len()
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
    if bytes.len() < 2 || !bytes.len().is_multiple_of(2) {
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
        // Project Unicode/alternate variants.
        0x0040 | 0x0041 | 0x003C | 0x0043 | 0x003D | 0x0042
            // Reference Unicode/alternate variants.
            | 0x003E
            // Module Unicode/alternate variants.
            | 0x0047 | 0x0032 | 0x0048 | 0x0049
    )
}

fn unicode_record_payload_len_prefixed(data: &[u8]) -> Option<&[u8]> {
    if data.len() < 4 {
        return None;
    }
    let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
    let remaining = data.len() - 4;
    if n == remaining || n.saturating_mul(2) == remaining {
        Some(&data[4..])
    } else {
        None
    }
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
        "-- ProjectNormalizedDataV3 (filtered PROJECT stream properties || V3ContentNormalizedData || FormsNormalizedData) --"
    );

    match formula_vba::project_normalized_data_v3_transcript(vba_project_bin) {
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
