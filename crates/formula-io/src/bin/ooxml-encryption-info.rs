use std::fs::File;
use std::io::{Read, Seek};
use std::path::PathBuf;

/// OLE/CFB file signature.
///
/// See: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

fn main() {
    let mut args = std::env::args_os();
    let exe = args
        .next()
        .unwrap_or_else(|| std::ffi::OsString::from("ooxml-encryption-info"));
    let Some(path) = args.next().map(PathBuf::from) else {
        eprintln!("usage: {} <path>", exe.to_string_lossy());
        std::process::exit(2);
    };
    if args.next().is_some() {
        eprintln!("usage: {} <path>", exe.to_string_lossy());
        std::process::exit(2);
    }

    let mut file = match File::open(&path) {
        Ok(f) => f,
        Err(err) => {
            eprintln!("error: failed to open {}: {err}", path.display());
            std::process::exit(1);
        }
    };

    let mut header = [0u8; 8];
    let n = match file.read(&mut header) {
        Ok(n) => n,
        Err(err) => {
            eprintln!("error: failed to read {}: {err}", path.display());
            std::process::exit(1);
        }
    };
    if n != OLE_MAGIC.len() || header != OLE_MAGIC {
        eprintln!(
            "not an OLE/CFB compound file (missing magic header): {}",
            path.display()
        );
        std::process::exit(1);
    }

    if let Err(err) = file.rewind() {
        eprintln!("error: failed to rewind {}: {err}", path.display());
        std::process::exit(1);
    }

    let mut ole = match cfb::CompoundFile::open(file) {
        Ok(ole) => ole,
        Err(err) => {
            eprintln!("error: failed to parse OLE/CFB compound file: {err}");
            std::process::exit(1);
        }
    };

    if !stream_exists(&mut ole, "EncryptionInfo") || !stream_exists(&mut ole, "EncryptedPackage") {
        eprintln!(
            "OLE/CFB file is not an OOXML encrypted container (missing EncryptionInfo/EncryptedPackage): {}",
            path.display()
        );
        std::process::exit(1);
    }

    let mut stream = match open_stream(&mut ole, "EncryptionInfo") {
        Ok(s) => s,
        Err(err) => {
            eprintln!("error: failed to open EncryptionInfo stream: {err}");
            std::process::exit(1);
        }
    };

    let mut info_hdr = [0u8; 8];
    if let Err(err) = stream.read_exact(&mut info_hdr) {
        eprintln!("error: failed to read EncryptionInfo header: {err}");
        std::process::exit(1);
    }
    let major = u16::from_le_bytes([info_hdr[0], info_hdr[1]]);
    let minor = u16::from_le_bytes([info_hdr[2], info_hdr[3]]);
    let flags = u32::from_le_bytes([info_hdr[4], info_hdr[5], info_hdr[6], info_hdr[7]]);

    // MS-OFFCRYPTO identifies "Standard" encryption via `versionMinor == 2`, but real-world files
    // vary the major version across Office generations (2/3/4). Keep this diagnostic tool aligned
    // with `formula-offcrypto` so it correctly labels Standard-encrypted files.
    let kind = match (major, minor) {
        (4, 4) => "Agile",
        (major, 2) if (2..=4).contains(&major) => "Standard",
        (major, 3) if (3..=4).contains(&major) => "Extensible",
        _ => "Unknown",
    };

    let mut extra = String::new();
    if kind == "Agile" {
        if let Some(tag) = sniff_agile_xml_root_tag(&mut stream) {
            extra.push_str(" xml_root=");
            extra.push_str(&tag);
        }
    }

    // One-line summary; keep deterministic to help with fixture validation and corpus triage.
    println!("{kind} ({major}.{minor}) flags=0x{flags:08x}{extra}");
}

fn stream_exists<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    ole.open_stream(&with_leading_slash).is_ok()
}

fn open_stream<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<cfb::Stream<R>, String> {
    match ole.open_stream(name) {
        Ok(s) => Ok(s),
        Err(err1) => {
            let with_leading_slash = format!("/{name}");
            match ole.open_stream(&with_leading_slash) {
                Ok(s) => Ok(s),
                Err(err2) => Err(format!("{err1}; {err2}")),
            }
        }
    }
}

fn sniff_agile_xml_root_tag<R: Read>(stream: &mut R) -> Option<String> {
    use std::io::Read as _;

    // The Agile EncryptionInfo stream payload is an XML document, typically small.
    // Read a bounded prefix to avoid accidentally loading huge/corrupt streams.
    const MAX_XML_PREFIX_BYTES: u64 = 128 * 1024; // 128KiB
    let mut buf = Vec::new();
    stream
        .take(MAX_XML_PREFIX_BYTES)
        .read_to_end(&mut buf)
        .ok()?;

    let xml = decode_best_effort_xml(&buf)?;
    xml_root_tag_name(&xml)
}

fn decode_best_effort_xml(buf: &[u8]) -> Option<String> {
    // UTF-8 BOM.
    let buf = buf.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(buf);

    // UTF-16 BOM.
    if let Some(rest) = buf.strip_prefix(&[0xFF, 0xFE]) {
        let (cow, _) = encoding_rs::UTF_16LE.decode_without_bom_handling(rest);
        return Some(cow.into_owned());
    }
    if let Some(rest) = buf.strip_prefix(&[0xFE, 0xFF]) {
        let (cow, _) = encoding_rs::UTF_16BE.decode_without_bom_handling(rest);
        return Some(cow.into_owned());
    }

    // If we see NUL bytes, best-effort guess UTF-16 endianness for ASCII-heavy XML.
    if buf.iter().any(|b| *b == 0) {
        let sample_len = buf.len().min(512);
        let sample_len = sample_len - (sample_len % 2);
        if sample_len >= 4 {
            let sample = &buf[..sample_len];
            let mut even_zero = 0usize;
            let mut odd_zero = 0usize;
            for (idx, b) in sample.iter().enumerate() {
                if *b == 0 {
                    if idx % 2 == 0 {
                        even_zero += 1;
                    } else {
                        odd_zero += 1;
                    }
                }
            }
            // For UTF-16LE ASCII, the high byte is typically 0, which lands at odd indexes.
            // For UTF-16BE ASCII, the high byte lands at even indexes.
            if odd_zero > even_zero.saturating_mul(3) {
                let (cow, _) = encoding_rs::UTF_16LE.decode_without_bom_handling(buf);
                return Some(cow.into_owned());
            } else if even_zero > odd_zero.saturating_mul(3) {
                let (cow, _) = encoding_rs::UTF_16BE.decode_without_bom_handling(buf);
                return Some(cow.into_owned());
            }
        }
    }

    std::str::from_utf8(buf).ok().map(|s| s.to_string())
}

fn xml_root_tag_name(mut xml: &str) -> Option<String> {
    xml = xml.trim_start();

    // Skip XML declaration / processing instructions / comments / doctypes until we find a start tag.
    loop {
        let start = xml.find('<')?;
        xml = &xml[start..];

        let after_lt = xml.get(1..)?;
        let mut chars = after_lt.chars();
        let first = chars.next()?;

        match first {
            '?' => {
                // Processing instruction (likely `<?xml ...?>`).
                let end = xml.find("?>")?;
                xml = xml.get(end + 2..)?.trim_start();
                continue;
            }
            '!' => {
                if xml.starts_with("<!--") {
                    let end = xml.find("-->")?;
                    xml = xml.get(end + 3..)?.trim_start();
                    continue;
                }
                // DOCTYPE or other declaration; skip to the next '>'.
                let end = xml.find('>')?;
                xml = xml.get(end + 1..)?.trim_start();
                continue;
            }
            '/' => {
                // Closing tag; skip.
                let end = xml.find('>')?;
                xml = xml.get(end + 1..)?.trim_start();
                continue;
            }
            _ => {
                // Start tag.
                let name_start = 1;
                let mut end_idx = name_start;
                for (i, ch) in xml[name_start..].char_indices() {
                    if ch.is_whitespace() || ch == '>' || ch == '/' {
                        break;
                    }
                    end_idx = name_start + i + ch.len_utf8();
                }
                if end_idx <= name_start {
                    return None;
                }
                return Some(xml[name_start..end_idx].to_string());
            }
        }
    }
}
