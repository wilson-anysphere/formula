//! Decrypt an OOXML `EncryptedPackage` container (password-protected `.xlsx` / `.xlsm` / `.xlsb`).
//!
//! Encrypted OOXML files are **not ZIP files on disk** even if they use a `.xlsx`/`.xlsm`/`.xlsb`
//! extension. Excel
//! wraps the real ZIP/OPC package in an OLE/CFB container with (at least) two streams:
//!
//! - `EncryptionInfo`
//! - `EncryptedPackage`
//!
//! This example reads those streams, prints a one-line `EncryptionInfo` summary to stderr, and
//! writes the decrypted ZIP bytes to a file or stdout.
//!
//! ## Usage
//!
//! ```bash
//! # Print help
//! cargo run -p formula-offcrypto --example decrypt_ooxml -- --help
//!
//! # Decrypt to a file
//! cargo run -p formula-offcrypto --example decrypt_ooxml -- \
//!   --input book.xlsx --password 'correct horse battery staple' --output book.zip
//!
//! # Decrypt to stdout (useful for piping)
//! cargo run -p formula-offcrypto --example decrypt_ooxml -- \
//!   --input book.xlsx --password 'pw' > book.zip
//!
//! # (Agile) Verify the `dataIntegrity` HMAC as well
//! cargo run -p formula-offcrypto --example decrypt_ooxml -- \
//!   --input book.xlsx --password 'pw' --verify-integrity > book.zip
//! ```
//!
//! The output is a ZIP file; you can inspect it with `unzip -l book.zip`.

use std::ffi::OsString;
use std::fs::File;
use std::io::{Read, Seek, Write};
use std::path::PathBuf;

use formula_offcrypto::{
    decrypt_encrypted_package, inspect_encryption_info, DecryptOptions,
};

fn main() {
    let args = match Args::parse() {
        Ok(args) => args,
        Err(ParseOutcome::Help(msg)) => {
            print!("{msg}");
            return;
        }
        Err(ParseOutcome::Error(msg)) => {
            eprintln!("{msg}");
            std::process::exit(2);
        }
    };

    let mut file = match File::open(&args.input) {
        Ok(f) => f,
        Err(err) => {
            eprintln!("error: failed to open {}: {err}", args.input.display());
            std::process::exit(1);
        }
    };

    let mut ole = match cfb::CompoundFile::open(&mut file) {
        Ok(ole) => ole,
        Err(err) => {
            eprintln!(
                "error: failed to parse OLE/CFB compound file {}: {err}",
                args.input.display()
            );
            std::process::exit(1);
        }
    };

    let encryption_info_bytes = match read_stream_best_effort(&mut ole, "EncryptionInfo") {
        Ok(b) => b,
        Err(err) => {
            eprintln!("error: failed to read EncryptionInfo stream: {err}");
            std::process::exit(1);
        }
    };
    let encrypted_package_bytes = match read_stream_best_effort(&mut ole, "EncryptedPackage") {
        Ok(b) => b,
        Err(err) => {
            eprintln!("error: failed to read EncryptedPackage stream: {err}");
            std::process::exit(1);
        }
    };

    match inspect_encryption_info(&encryption_info_bytes) {
        Ok(summary) => eprintln!("EncryptionInfo: {summary:?}"),
        Err(err) => eprintln!("warning: failed to inspect EncryptionInfo: {err}"),
    }

    let options = DecryptOptions {
        verify_integrity: args.verify_integrity,
        ..Default::default()
    };
    let decrypted_zip = match decrypt_encrypted_package(
        &encryption_info_bytes,
        &encrypted_package_bytes,
        &args.password,
        options,
    ) {
        Ok(b) => b,
        Err(err) => {
            eprintln!("error: failed to decrypt EncryptedPackage: {err}");
            std::process::exit(1);
        }
    };

    if let Some(out_path) = &args.output {
        if let Err(err) = std::fs::write(out_path, &decrypted_zip) {
            eprintln!("error: failed to write {}: {err}", out_path.display());
            std::process::exit(1);
        }
    } else {
        let mut stdout = std::io::stdout().lock();
        if let Err(err) = stdout.write_all(&decrypted_zip) {
            eprintln!("error: failed to write decrypted bytes to stdout: {err}");
            std::process::exit(1);
        }
    }
}

struct Args {
    input: PathBuf,
    password: String,
    verify_integrity: bool,
    output: Option<PathBuf>,
}

enum ParseOutcome {
    Help(String),
    Error(String),
}

impl Args {
    fn parse() -> Result<Self, ParseOutcome> {
        let mut input: Option<PathBuf> = None;
        let mut password: Option<String> = None;
        let mut verify_integrity = false;
        let mut output: Option<PathBuf> = None;

        let mut argv = std::env::args_os();
        let exe = argv
            .next()
            .unwrap_or_else(|| OsString::from("decrypt_ooxml"));

        while let Some(arg) = argv.next() {
            match arg.to_string_lossy().as_ref() {
                "-h" | "--help" => {
                    return Err(ParseOutcome::Help(Self::help(&exe)));
                }
                "--input" => {
                    let Some(v) = argv.next() else {
                        return Err(ParseOutcome::Error(format!(
                            "error: --input requires a value\n\n{}",
                            Self::help(&exe)
                        )));
                    };
                    input = Some(PathBuf::from(v));
                }
                "--password" => {
                    let Some(v) = argv.next() else {
                        return Err(ParseOutcome::Error(format!(
                            "error: --password requires a value\n\n{}",
                            Self::help(&exe)
                        )));
                    };
                    password = Some(v.to_string_lossy().to_string());
                }
                "--verify-integrity" => {
                    verify_integrity = true;
                }
                "--output" => {
                    let Some(v) = argv.next() else {
                        return Err(ParseOutcome::Error(format!(
                            "error: --output requires a value\n\n{}",
                            Self::help(&exe)
                        )));
                    };
                    output = Some(PathBuf::from(v));
                }
                other => {
                    return Err(ParseOutcome::Error(format!(
                        "error: unrecognized argument `{other}`\n\n{}",
                        Self::help(&exe)
                    )));
                }
            }
        }

        let input = input.ok_or_else(|| {
            ParseOutcome::Error(format!(
                "error: missing required --input\n\n{}",
                Self::help(&exe)
            ))
        })?;
        let password = password.ok_or_else(|| {
            ParseOutcome::Error(format!(
                "error: missing required --password\n\n{}",
                Self::help(&exe)
            ))
        })?;

        Ok(Self {
            input,
            password,
            verify_integrity,
            output,
        })
    }

    fn help(exe: &OsString) -> String {
        let exe = exe.to_string_lossy();
        format!(
            "Usage: {exe} --input <path> --password <pw> [--verify-integrity] [--output <path>]\n\
             \n\
             Decrypt an OOXML encrypted container (OLE/CFB with EncryptionInfo + EncryptedPackage).\n\
             \n\
             Options:\n\
               --input <path>           Path to the encrypted OLE/CFB file (.xlsx/.xlsm/.xlsb)\n\
               --password <pw>          Password to open the workbook\n\
               --verify-integrity       (Agile) verify dataIntegrity HMAC\n\
               --output <path>          Write decrypted ZIP bytes to a file (defaults to stdout)\n\
               -h, --help               Print help\n"
        )
    }
}

fn read_stream_best_effort<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<Vec<u8>, std::io::Error> {
    let mut stream = open_stream_best_effort(ole, name)?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf)?;
    Ok(buf)
}

fn open_stream_best_effort<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<cfb::Stream<R>, std::io::Error> {
    let want = name.trim_start_matches('/');

    if let Ok(s) = ole.open_stream(want) {
        return Ok(s);
    }
    let with_leading_slash = format!("/{want}");
    if let Ok(s) = ole.open_stream(&with_leading_slash) {
        return Ok(s);
    }

    // Case-insensitive fallback: walk the directory tree and match stream paths.
    let mut found_path: Option<String> = None;
    let mut found_normalized: Option<String> = None;
    for entry in ole.walk() {
        if !entry.is_stream() {
            continue;
        }
        let path = entry.path().to_string_lossy().to_string();
        let normalized = path.trim_start_matches('/').to_string();
        if normalized.eq_ignore_ascii_case(want) {
            found_path = Some(path);
            found_normalized = Some(normalized);
            break;
        }
    }

    if let Some(normalized) = found_normalized {
        if let Ok(s) = ole.open_stream(&normalized) {
            return Ok(s);
        }
        let with_slash = format!("/{normalized}");
        if let Ok(s) = ole.open_stream(&with_slash) {
            return Ok(s);
        }
        if let Some(path) = found_path {
            if let Ok(s) = ole.open_stream(&path) {
                return Ok(s);
            }
        }
    }

    Err(std::io::Error::new(
        std::io::ErrorKind::NotFound,
        format!("stream not found: `{want}`"),
    ))
}
