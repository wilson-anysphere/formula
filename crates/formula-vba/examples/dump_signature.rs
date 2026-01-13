use std::ffi::OsString;
use std::path::PathBuf;
use std::process::ExitCode;

use formula_vba::{
    extract_signer_certificate_info, extract_vba_signature_signed_digest,
    list_vba_digital_signatures, verify_vba_digital_signature,
};

#[path = "shared/vba_project_bin.rs"]
mod vba_project_bin;

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

    let (input_path, password) = parse_args(&program, args)?;

    let (vba_project_bin, source) =
        vba_project_bin::load_vba_project_bin(&input_path, password.as_deref())?;

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
        "usage: {} [--password <pw>] <vbaProject.bin|workbook.xlsm|workbook.xlsx|workbook.xlsb>",
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

    while let Some(arg) = args.next() {
        let arg_str = arg.to_string_lossy();

        if arg_str == "--help" || arg_str == "-h" {
            return Err(usage(program));
        }

        if arg_str == "--password" {
            let Some(value) = args.next() else {
                return Err(format!("missing value for --password\n{}", usage(program)));
            };
            let value = value
                .to_str()
                .ok_or_else(|| "--password value must be valid UTF-8".to_owned())?
                .to_owned();
            password = Some(value);
            continue;
        }

        if let Some(rest) = arg_str.strip_prefix("--password=") {
            password = Some(rest.to_owned());
            continue;
        }

        if input.is_none() {
            input = Some(PathBuf::from(arg));
            continue;
        }

        return Err(usage(program));
    }

    let Some(input) = input else {
        return Err(usage(program));
    };
    Ok((input, password))
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
