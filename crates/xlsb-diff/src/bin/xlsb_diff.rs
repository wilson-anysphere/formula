fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = xlsx_diff::cli::parse_args();
    eprintln!("warning: `xlsb-diff` is deprecated; use `xlsx-diff` (it supports .xlsb)");
    xlsx_diff::cli::run_with_args(args).map_err(|err| err.into())
}
