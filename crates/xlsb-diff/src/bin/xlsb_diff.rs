fn main() -> Result<(), Box<dyn std::error::Error>> {
    eprintln!(
        "warning: `xlsb-diff` is deprecated; use `xlsx-diff` (it supports .xlsb and encrypted workbooks via --password/--password-file)"
    );
    let args = xlsx_diff::cli::parse_args();
    xlsx_diff::cli::run_with_args(args).map_err(|err| err.into())
}
