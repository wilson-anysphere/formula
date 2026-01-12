use std::ffi::OsString;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use formula_vba::{
    extract_signer_certificate_info, extract_vba_signature_signed_digest,
    list_vba_digital_signatures, verify_vba_digital_signature,
};

#[path = "shared/zip_util.rs"]
mod zip_util;

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
        .unwrap_or_else(|| OsString::from("dump_signature"));

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

    let signatures = list_vba_digital_signatures(&vba_project_bin)
        .map_err(|e| format!("failed to list VBA signature streams: {e}"))?;

    if signatures.is_empty() {
        println!("signature streams: none");
    } else {
        println!("signature streams: {}", signatures.len());
        for (idx, sig) in signatures.iter().enumerate() {
            println!();
            println!(
                "[{}] stream_path: {}",
                idx + 1,
                escape_ole_path(&sig.stream_path)
            );
            println!("    stream_len: {} bytes", sig.signature.len());
            println!("    pkcs7_verification: {:?}", sig.verification);
            println!(
                "    signer_subject: {}",
                sig.signer_subject.as_deref().unwrap_or("<not found>")
            );

            match extract_signer_certificate_info(&sig.signature) {
                Some(info) => {
                    println!("    signer_cert_subject: {}", info.subject);
                    println!("    signer_cert_issuer: {}", info.issuer);
                    println!("    signer_cert_serial: {}", info.serial_hex);
                    println!(
                        "    signer_cert_sha256_fingerprint: {}",
                        info.sha256_fingerprint_hex
                    );
                }
                None => {
                    println!("    signer_cert_subject: <not found>");
                    println!("    signer_cert_issuer: <not found>");
                    println!("    signer_cert_sha256_fingerprint: <not found>");
                }
            }

            match (sig.pkcs7_offset, sig.pkcs7_len) {
                (Some(offset), Some(len)) => {
                    println!("    pkcs7_location: offset={offset} len={len} (DigSig wrapper)");
                }
                _ => {
                    println!("    pkcs7_location: <unknown> (no DigSig wrapper detected)");
                }
            }
            if let Some(version) = sig.digsig_info_version {
                println!("    digsig_info_version: {version}");
            }

            if let (Some(oid), Some(digest)) = (
                sig.signed_digest_algorithm_oid.as_deref(),
                sig.signed_digest.as_deref(),
            ) {
                println!(
                    "    signed_digest: alg_oid={oid} digest={}",
                    bytes_to_lower_hex(digest)
                );
            } else {
                match extract_vba_signature_signed_digest(&sig.signature) {
                    Ok(Some(digest)) => {
                        println!(
                            "    signed_digest: alg_oid={} digest={}",
                            digest.digest_algorithm_oid,
                            bytes_to_lower_hex(&digest.digest)
                        );
                    }
                    Ok(None) => {
                        println!("    signed_digest: <not found>");
                    }
                    Err(err) => {
                        // Keep going: this tool is intended for debugging partially malformed inputs.
                        println!("    signed_digest: <error: {err}>");
                    }
                }
            }
        }
    }

    println!();
    println!("verify_vba_digital_signature:");
    match verify_vba_digital_signature(&vba_project_bin)
        .map_err(|e| format!("failed to verify VBA digital signature: {e}"))?
    {
        None => println!("    signature: none"),
        Some(sig) => {
            println!("    chosen_stream: {}", escape_ole_path(&sig.stream_path));
            println!("    pkcs7_verification: {:?}", sig.verification);
            println!("    binding: {:?}", sig.binding);
        }
    }

    Ok(())
}

fn usage(program: &OsString) -> String {
    format!(
        "usage: {} <vbaProject.bin|workbook.xlsm|workbook.xlsx>",
        program.to_string_lossy()
    )
}

fn load_vba_project_bin(path: &Path) -> Result<(Vec<u8>, String), String> {
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    if ext == "xlsm" || ext == "xlsx" {
        let bytes = extract_vba_project_bin_from_zip(path)?;
        return Ok((
            bytes,
            format!("{} (zip entry xl/vbaProject.bin)", path.display()),
        ));
    }

    // Treat as a raw vbaProject.bin OLE file.
    let bytes =
        std::fs::read(path).map_err(|e| format!("failed to read {}: {e}", path.display()))?;
    Ok((bytes, path.display().to_string()))
}

fn extract_vba_project_bin_from_zip(path: &Path) -> Result<Vec<u8>, String> {
    let file = match std::fs::File::open(path) {
        Ok(f) => f,
        Err(e) => return Err(format!("failed to open {}: {e}", path.display())),
    };

    let mut archive = match zip::ZipArchive::new(file) {
        Ok(a) => a,
        Err(e) => return Err(format!("failed to open zip {}: {e}", path.display())),
    };

    let Some(buf) = zip_util::read_zip_entry_bytes(&mut archive, "xl/vbaProject.bin")
        .map_err(|e| format!("failed to read zip {}: {e}", path.display()))?
    else {
        return Err(format!(
            "{} is a zip, but does not contain xl/vbaProject.bin",
            path.display()
        ));
    };
    Ok(buf)
}

fn escape_ole_path(path: &str) -> String {
    path.chars()
        .flat_map(|c| c.escape_default())
        .collect::<String>()
}

fn bytes_to_lower_hex(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut out = String::with_capacity(bytes.len() * 2);
    for &b in bytes {
        out.push(HEX[(b >> 4) as usize] as char);
        out.push(HEX[(b & 0x0f) as usize] as char);
    }
    out
}
