#!/usr/bin/env python3

from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import platform
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

DEFAULT_TOP = 50
DEFAULT_FEATURES = "desktop"
DEFAULT_BIN_NAME = "formula-desktop"
DEFAULT_DESKTOP_PKG_FALLBACK = "formula-desktop-tauri"
ENV_SIZE_LIMIT_MB = "FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB"
ENV_ENFORCE_SIZE_LIMIT = "FORMULA_ENFORCE_DESKTOP_BINARY_SIZE"


@dataclass(frozen=True)
class CmdResult:
    cmd: list[str]
    returncode: int
    stdout: str
    stderr: str

    @property
    def combined(self) -> str:
        out = self.stdout or ""
        err = self.stderr or ""
        if err.strip():
            if out.strip():
                return f"{out.rstrip()}\n{err.rstrip()}"
            return err.rstrip()
        return out.rstrip()


def _human_bytes(size_bytes: int) -> str:
    size = float(size_bytes)
    units = ["B", "KB", "MB", "GB", "TB"]
    for unit in units:
        if size < 1000 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}"
        size /= 1000
    return f"{size_bytes} B"


def _write_json_out(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(payload, f, indent=2, sort_keys=True)
        f.write("\n")


def _is_truthy_env(val: str | None) -> bool:
    if val is None:
        return False
    return val.strip().lower() in {"1", "true", "yes", "y", "on"}


def _parse_limit_mb(raw: str | None) -> int | None:
    if raw is None or not raw.strip():
        return None
    try:
        mb = int(float(raw))
    except ValueError as exc:  # noqa: PERF203
        raise ValueError(f"Invalid {ENV_SIZE_LIMIT_MB}={raw!r} (expected a number)") from exc
    if mb <= 0:
        raise ValueError(f"Invalid {ENV_SIZE_LIMIT_MB}={raw!r} (must be > 0)")
    return mb


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _relpath(path: Path, repo_root: Path) -> str:
    try:
        return str(path.relative_to(repo_root))
    except ValueError:
        return str(path)


def _desktop_package_name(repo_root: Path) -> str:
    """
    The desktop Tauri shell has historically used both `desktop` and
    `formula-desktop-tauri` as the Cargo package name. Prefer reading the current
    name from `apps/desktop/src-tauri/Cargo.toml` to keep scripts stable.
    """
    cargo_toml = repo_root / "apps" / "desktop" / "src-tauri" / "Cargo.toml"
    if not cargo_toml.is_file():
        return DEFAULT_DESKTOP_PKG_FALLBACK

    in_package = False
    for raw in cargo_toml.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("[") and line.endswith("]"):
            in_package = line == "[package]"
            continue
        if not in_package:
            continue
        if line.startswith("name"):
            # Accept `name = "..."` with arbitrary whitespace.
            parts = line.split("=", 1)
            if len(parts) != 2:
                continue
            rhs = parts[1].strip()
            if rhs.startswith('"') and '"' in rhs[1:]:
                return rhs.split('"', 2)[1]
    return DEFAULT_DESKTOP_PKG_FALLBACK


def _cargo_target_directory(repo_root: Path) -> Path:
    cp = subprocess.run(
        ["cargo", "metadata", "--no-deps", "--format-version=1"],
        cwd=repo_root,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    meta = json.loads(cp.stdout)
    target_dir = meta.get("target_directory")
    if not target_dir:
        raise RuntimeError("cargo metadata did not return target_directory")
    return Path(target_dir)


def _binary_path(target_dir: Path, bin_name: str, target: str | None) -> Path:
    rel_dir = target_dir / "release"
    if target:
        rel_dir = target_dir / target / "release"
    exe = bin_name
    # If cross-compiling to Windows from a non-Windows host, the output binary
    # still has a `.exe` suffix under target/<triple>/release.
    if (target and "windows" in target) or (target is None and sys.platform == "win32"):
        exe += ".exe"
    return rel_dir / exe


def _candidate_binary_paths(repo_root: Path, target_dir: Path, bin_name: str, target: str | None) -> list[Path]:
    """
    Return a small set of plausible binary paths.

    In this repo, workspace builds typically land in `<repo>/target`, but some
    Tauri workflows (or historical layouts) may place artifacts under
    `apps/desktop/src-tauri/target`.
    """
    candidate_target_dirs = [
        target_dir,
        repo_root / "target",
        repo_root / "apps" / "desktop" / "src-tauri" / "target",
    ]

    # De-dupe while preserving order.
    seen: set[Path] = set()
    uniq_target_dirs: list[Path] = []
    for td in candidate_target_dirs:
        try:
            key = td.resolve()
        except OSError:
            key = td
        if key in seen:
            continue
        seen.add(key)
        uniq_target_dirs.append(td)

    candidates: list[Path] = []
    for td in uniq_target_dirs:
        candidates.append(_binary_path(td, bin_name, target))
        # If callers pass --target but the binary was built for the host
        # (or vice-versa), check the alternate location too.
        if target is not None:
            candidates.append(_binary_path(td, bin_name, None))

    # De-dupe paths.
    seen_paths: set[Path] = set()
    uniq_paths: list[Path] = []
    for p in candidates:
        try:
            key = p.resolve()
        except OSError:
            key = p
        if key in seen_paths:
            continue
        seen_paths.add(key)
        uniq_paths.append(p)
    return uniq_paths


def _first_existing_path(paths: list[Path]) -> Path | None:
    for p in paths:
        if p.is_file():
            return p
    return None


def _run_capture(cmd: list[str], *, cwd: Path) -> CmdResult:
    cp = subprocess.run(
        cmd,
        cwd=cwd,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return CmdResult(cmd=cmd, returncode=cp.returncode, stdout=cp.stdout, stderr=cp.stderr)


def _append_step_summary(markdown: str) -> None:
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not summary_path:
        return
    try:
        with open(summary_path, "a", encoding="utf-8", errors="replace") as f:
            f.write(markdown)
            f.write("\n")
    except OSError:
        # Don't fail the report if GitHub step summary can't be written.
        return


def _render_cmd(cmd: list[str]) -> str:
    # Use shell-ish quoting for readability.
    out: list[str] = []
    for part in cmd:
        if any(c in part for c in (" ", "\t", "\n", '"', "'", "$")):
            out.append(json.dumps(part))
        else:
            out.append(part)
    return " ".join(out)


def _render_markdown(
    *,
    package: str,
    bin_name: str,
    features: str,
    target: str | None,
    repo_root: Path,
    target_dir: Path,
    bin_path: Path,
    bin_size_bytes: int | None,
    limit_mb: int | None,
    enforce: bool,
    rustc_version: str | None,
    cargo_version: str | None,
    git_sha: str | None,
    git_ref: str | None,
    cargo_bloat_version: str | None,
    file_info: str | None,
    stripped: bool | None,
    build_ran: bool,
    json_out_path: str | None,
    build_cmd: list[str] | None,
    crates_cmd: list[str] | None,
    crates_out: CmdResult | None,
    symbols_cmd: list[str] | None,
    symbols_out: CmdResult | None,
    llvm_size_cmd: list[str] | None,
    llvm_size_out: CmdResult | None,
    tool_note: str | None,
) -> str:
    lines: list[str] = []
    runner_os = os.environ.get("RUNNER_OS", "").strip() or platform.system()

    heading = "## Desktop Rust binary size breakdown"
    if runner_os:
        heading += f" ({runner_os})"
    lines.append(heading)
    lines.append("")

    lines.append(f"- Package: `{package}`")
    lines.append(f"- Binary: `{bin_name}` (profile: `release`)")
    lines.append(f"- Features: `{features}`")
    if target:
        lines.append(f"- Target: `{target}`")
    if rustc_version:
        lines.append(f"- rustc: `{rustc_version}`")
    if cargo_version:
        lines.append(f"- cargo: `{cargo_version}`")
    if git_sha:
        lines.append(f"- Git SHA: `{git_sha}`")
    if git_ref:
        lines.append(f"- Git ref: `{git_ref}`")
    lines.append(f"- Target dir: `{_relpath(target_dir, repo_root)}`")
    lines.append(f"- Binary path: `{_relpath(bin_path, repo_root)}`")
    if bin_size_bytes is not None:
        lines.append(f"- Binary size: **{_human_bytes(bin_size_bytes)}** ({bin_size_bytes} bytes)")
    if file_info:
        lines.append(f"- File info: `{file_info}`")
    if stripped is not None:
        lines.append(f"- Stripped: **{'yes' if stripped else 'no'}**")
    if cargo_bloat_version:
        lines.append(f"- cargo-bloat: `{cargo_bloat_version}`")
    if limit_mb is None:
        lines.append(f"- Size limit: _(not set)_ (set `{ENV_SIZE_LIMIT_MB}=...` to enable)")
        lines.append(
            f"- Enforcement: **{'enabled' if enforce else 'disabled'}** "
            f"(set `{ENV_ENFORCE_SIZE_LIMIT}=1` to fail when over the limit)"
        )
    else:
        over_limit = bin_size_bytes is not None and bin_size_bytes > (limit_mb * 1_000_000)
        lines.append(f"- Size limit: **{limit_mb} MB**")
        lines.append(
            f"- Enforcement: **{'enabled' if enforce else 'disabled'}** "
            f"(set `{ENV_ENFORCE_SIZE_LIMIT}=1` to fail when over the limit)"
        )
        if bin_size_bytes is not None:
            lines.append(f"- Over limit: **{'YES' if over_limit else 'no'}**")
    if json_out_path:
        lines.append(f"- JSON report: `{json_out_path}`")
    lines.append("")

    if tool_note:
        lines.append(tool_note)
        lines.append("")

    if build_cmd:
        lines.append("### Build command")
        lines.append("")
        if not build_ran:
            lines.append("_(not executed; this report was generated with `--no-build`)_")
            lines.append("")
        lines.append("```bash")
        lines.append(_render_cmd(build_cmd))
        lines.append("```")
        lines.append("")

    if crates_cmd and crates_out:
        lines.append("### Top crates (cargo-bloat)")
        lines.append("")
        lines.append("Command:")
        lines.append(f"`{_render_cmd(crates_cmd)}`")
        lines.append("")
        lines.append("```text")
        lines.append(crates_out.combined or "<no output>")
        lines.append("```")
        lines.append("")

    if symbols_cmd and symbols_out:
        lines.append("### Top symbols (cargo-bloat)")
        lines.append("")
        lines.append("Command:")
        lines.append(f"`{_render_cmd(symbols_cmd)}`")
        lines.append("")
        lines.append("```text")
        lines.append(symbols_out.combined or "<no output>")
        lines.append("```")
        lines.append("")

    if llvm_size_cmd and llvm_size_out:
        tool_label = Path(llvm_size_cmd[0]).name
        lines.append(f"### Section sizes ({tool_label})")
        lines.append("")
        lines.append("Command:")
        lines.append(f"`{_render_cmd(llvm_size_cmd)}`")
        lines.append("")
        lines.append("```text")
        lines.append(llvm_size_out.combined or "<no output>")
        lines.append("```")
        lines.append("")

    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate a binary size breakdown report for the desktop Rust/Tauri shell."
    )
    parser.add_argument(
        "--no-build",
        action="store_true",
        default=False,
        help="Assume the release binary is already built and skip `cargo build`.",
    )
    parser.add_argument(
        "--top",
        type=int,
        default=DEFAULT_TOP,
        help=f"How many entries to show from cargo-bloat (default: {DEFAULT_TOP}).",
    )
    parser.add_argument(
        "--target",
        type=str,
        default=None,
        help="Optional Cargo target triple (e.g. x86_64-unknown-linux-gnu).",
    )
    parser.add_argument(
        "--features",
        type=str,
        default=DEFAULT_FEATURES,
        help=f"Cargo feature set to enable (default: {DEFAULT_FEATURES}).",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Optional markdown output path (in addition to stdout).",
    )
    parser.add_argument(
        "--json-out",
        type=Path,
        default=None,
        help="Optional path to write a machine-readable JSON report.",
    )
    parser.add_argument(
        "--limit-mb",
        type=int,
        default=None,
        help=(
            "Optional absolute size limit (in MB) for the release binary. "
            f"Also configurable via env {ENV_SIZE_LIMIT_MB}."
        ),
    )
    parser.add_argument(
        "--enforce",
        action="store_true",
        default=False,
        help=(
            "Fail if the release binary exceeds --limit-mb (or env limit). "
            f"Also enabled by env {ENV_ENFORCE_SIZE_LIMIT}=1."
        ),
    )
    args = parser.parse_args()

    repo_root = _repo_root()
    # `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml`. Some environments set it
    # globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift
    # for any cargo/rustc subprocess calls in this report.
    if os.environ.get("RUSTUP_TOOLCHAIN") and (repo_root / "rust-toolchain.toml").is_file():
        os.environ.pop("RUSTUP_TOOLCHAIN", None)

    # Use a repo-local cargo home by default to avoid lock contention on ~/.cargo
    # when many agents build/test concurrently. Preserve any explicit override and
    # keep CI caching behavior intact.
    #
    # Note: some environments pre-set `CARGO_HOME=$HOME/.cargo`. Treat that value as
    # "unset" in non-CI runs so we still get per-repo isolation by default.
    # To explicitly keep `CARGO_HOME=$HOME/.cargo` in local runs, set
    # `FORMULA_ALLOW_GLOBAL_CARGO_HOME=1`.
    default_global_cargo_home = Path.home() / ".cargo"
    cargo_home = os.environ.get("CARGO_HOME")
    cargo_home_path = Path(cargo_home).expanduser() if cargo_home else None
    is_ci = _is_truthy_env(os.environ.get("CI"))
    if not cargo_home or (
        not is_ci
        and not os.environ.get("FORMULA_ALLOW_GLOBAL_CARGO_HOME")
        and cargo_home_path == default_global_cargo_home
    ):
        os.environ["CARGO_HOME"] = str(repo_root / "target" / "cargo-home")
    Path(os.environ["CARGO_HOME"]).mkdir(parents=True, exist_ok=True)

    # Ensure tools installed via `cargo install` under this CARGO_HOME are available.
    cargo_bin_dir = Path(os.environ["CARGO_HOME"]) / "bin"
    cargo_bin_dir.mkdir(parents=True, exist_ok=True)
    path_entries = os.environ.get("PATH", "").split(os.pathsep) if os.environ.get("PATH") else []
    if str(cargo_bin_dir) not in path_entries:
        os.environ["PATH"] = (
            f"{cargo_bin_dir}{os.pathsep}{os.environ['PATH']}" if os.environ.get("PATH") else str(cargo_bin_dir)
        )

    package = _desktop_package_name(repo_root)
    bin_name = DEFAULT_BIN_NAME
    features = args.features
    target = args.target
    json_out_path = _relpath(args.json_out, repo_root) if args.json_out else None
    enforce = args.enforce or _is_truthy_env(os.environ.get(ENV_ENFORCE_SIZE_LIMIT))
    try:
        limit_mb = args.limit_mb if args.limit_mb is not None else _parse_limit_mb(os.environ.get(ENV_SIZE_LIMIT_MB))
    except ValueError as exc:
        md = f"## Desktop Rust binary size breakdown\n\nError: {exc}\n"
        print(md)
        _append_step_summary(md)
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(md, encoding="utf-8")
        return 2

    if enforce and limit_mb is None:
        md = (
            "## Desktop Rust binary size breakdown\n\n"
            f"Error: {ENV_ENFORCE_SIZE_LIMIT} is enabled but no size limit was provided. "
            f"Set --limit-mb or {ENV_SIZE_LIMIT_MB}.\n"
        )
        print(md)
        _append_step_summary(md)
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(md, encoding="utf-8")
        return 2

    rustc_version: str | None = None
    cargo_version: str | None = None
    git_sha: str | None = None
    git_ref: str | None = None
    try:
        rustc = _run_capture(["rustc", "--version"], cwd=repo_root)
        if rustc.returncode == 0:
            rustc_version = rustc.stdout.strip().splitlines()[0] if rustc.stdout.strip() else None
    except FileNotFoundError:
        rustc_version = None
    try:
        cargo_ver = _run_capture(["cargo", "--version"], cwd=repo_root)
        if cargo_ver.returncode == 0:
            cargo_version = cargo_ver.stdout.strip().splitlines()[0] if cargo_ver.stdout.strip() else None
    except FileNotFoundError:
        cargo_version = None

    # Record git metadata for reproducibility.
    git_ref = os.environ.get("GITHUB_REF_NAME") or os.environ.get("GITHUB_REF")
    git_sha = os.environ.get("GITHUB_SHA")
    if git_sha:
        git_sha = git_sha.strip()
        if len(git_sha) > 12:
            git_sha = git_sha[:12]
    else:
        try:
            rev = _run_capture(["git", "rev-parse", "--short=12", "HEAD"], cwd=repo_root)
            if rev.returncode == 0:
                git_sha = rev.stdout.strip().splitlines()[0] if rev.stdout.strip() else None
        except FileNotFoundError:
            git_sha = None

    try:
        target_dir = _cargo_target_directory(repo_root)
    except Exception as exc:  # noqa: BLE001
        md = f"## Desktop Rust binary size breakdown\n\nError: failed to run `cargo metadata`: `{exc}`\n"
        print(md)
        _append_step_summary(md)
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(md, encoding="utf-8")
        if args.json_out:
            _write_json_out(
                args.json_out,
                {
                    "status": "error",
                    "error": f"cargo metadata failed: {exc}",
                    "generated_at": _dt.datetime.now(tz=_dt.timezone.utc).isoformat(),
                    "runner": {
                        "os": os.environ.get("RUNNER_OS") or platform.system(),
                        "arch": os.environ.get("RUNNER_ARCH") or platform.machine(),
                    },
                    "git": {"sha": git_sha, "ref": git_ref},
                    "toolchain": {"rustc": rustc_version, "cargo": cargo_version},
                    "package": package,
                    "bin_name": bin_name,
                    "features": features,
                    "target": target,
                },
            )
        return 1

    candidate_bin_paths = _candidate_binary_paths(repo_root, target_dir, bin_name, target)
    default_bin_path = candidate_bin_paths[0] if candidate_bin_paths else _binary_path(target_dir, bin_name, target)
    bin_path = default_bin_path

    build_cmd: list[str] | None = [
        "cargo",
        "build",
        "-p",
        package,
        "--bin",
        bin_name,
        "--features",
        features,
        "--release",
        "--locked",
    ]
    if target:
        build_cmd.extend(["--target", target])

    build_ran = False
    if not args.no_build:
        build_ran = True
        try:
            subprocess.run(build_cmd, cwd=repo_root, check=True)
        except subprocess.CalledProcessError:
            md = _render_markdown(
                package=package,
                bin_name=bin_name,
                features=features,
                target=target,
                repo_root=repo_root,
                target_dir=target_dir,
                bin_path=default_bin_path,
                bin_size_bytes=bin_path.stat().st_size if bin_path.exists() else None,
                limit_mb=limit_mb,
                enforce=enforce,
                rustc_version=rustc_version,
                cargo_version=cargo_version,
                git_sha=git_sha,
                git_ref=git_ref,
                cargo_bloat_version=None,
                file_info=None,
                stripped=None,
                build_ran=build_ran,
                json_out_path=json_out_path,
                build_cmd=build_cmd,
                crates_cmd=None,
                crates_out=None,
                symbols_cmd=None,
                symbols_out=None,
                llvm_size_cmd=None,
                llvm_size_out=None,
                tool_note="Build failed (see logs above). The size breakdown report could not be generated.",
            )
            print(md)
            _append_step_summary(md)
            if args.out:
                args.out.parent.mkdir(parents=True, exist_ok=True)
                args.out.write_text(md, encoding="utf-8")
            if args.json_out:
                _write_json_out(
                    args.json_out,
                    {
                        "status": "error",
                        "error": "build failed",
                        "generated_at": _dt.datetime.now(tz=_dt.timezone.utc).isoformat(),
                        "runner": {"os": os.environ.get("RUNNER_OS") or platform.system(), "arch": os.environ.get("RUNNER_ARCH") or platform.machine()},
                        "git": {"sha": git_sha, "ref": git_ref},
                        "package": package,
                        "bin_name": bin_name,
                        "features": features,
                        "target": target,
                        "toolchain": {"rustc": rustc_version, "cargo": cargo_version},
                        "build_cmd": build_cmd,
                        "build_ran": build_ran,
                    },
                )
            return 1

    existing = _first_existing_path(candidate_bin_paths)
    if existing is not None:
        bin_path = existing

    if not bin_path.exists():
        searched_lines = "\n".join(f"- `{_relpath(p, repo_root)}`" for p in candidate_bin_paths[:8])
        if len(candidate_bin_paths) > 8:
            searched_lines += "\n- …"
        md = _render_markdown(
            package=package,
            bin_name=bin_name,
            features=features,
            target=target,
            repo_root=repo_root,
            target_dir=target_dir,
            bin_path=default_bin_path,
            bin_size_bytes=None,
            limit_mb=limit_mb,
            enforce=enforce,
            rustc_version=rustc_version,
            cargo_version=cargo_version,
            git_sha=git_sha,
            git_ref=git_ref,
            cargo_bloat_version=None,
            file_info=None,
            stripped=None,
            build_ran=build_ran,
            json_out_path=json_out_path,
            build_cmd=build_cmd,
            crates_cmd=None,
            crates_out=None,
            symbols_cmd=None,
            symbols_out=None,
            llvm_size_cmd=None,
            llvm_size_out=None,
            tool_note=(
                "Binary not found.\n\n"
                "Searched:\n"
                f"{searched_lines}\n\n"
                "Hint: run the build command in this report (or omit `--no-build`)."
            ),
        )
        print(md)
        _append_step_summary(md)
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(md, encoding="utf-8")
        if args.json_out:
            _write_json_out(
                args.json_out,
                {
                    "status": "error",
                    "error": "binary not found",
                    "generated_at": _dt.datetime.now(tz=_dt.timezone.utc).isoformat(),
                    "runner": {"os": os.environ.get("RUNNER_OS") or platform.system(), "arch": os.environ.get("RUNNER_ARCH") or platform.machine()},
                    "git": {"sha": git_sha, "ref": git_ref},
                    "package": package,
                    "bin_name": bin_name,
                    "features": features,
                    "target": target,
                    "toolchain": {"rustc": rustc_version, "cargo": cargo_version},
                    "target_dir": _relpath(target_dir, repo_root),
                    "bin_path": _relpath(default_bin_path, repo_root),
                    "searched_paths": [_relpath(p, repo_root) for p in candidate_bin_paths],
                    "build_cmd": build_cmd,
                    "build_ran": build_ran,
                },
            )
        return 1

    bin_size_bytes = bin_path.stat().st_size
    limit_bytes = limit_mb * 1000 * 1000 if limit_mb is not None else None
    over_limit = limit_bytes is not None and bin_size_bytes > limit_bytes

    # Best-effort: show `file` output (helps spot unstripped/debug builds).
    file_info: str | None = None
    stripped: bool | None = None
    file_tool = shutil.which("file")
    if file_tool:
        file_out = _run_capture([file_tool, str(bin_path)], cwd=repo_root)
        if file_out.returncode == 0:
            line = file_out.stdout.strip().splitlines()[0] if file_out.stdout.strip() else ""
            if line:
                _prefix, sep, rest = line.partition(":")
                file_info = rest.strip() if sep else line.strip()
                lowered = file_info.lower()
                if "not stripped" in lowered:
                    stripped = False
                elif "stripped" in lowered:
                    stripped = True

    # Preferred: cargo-bloat.
    cargo_bloat_probe = _run_capture(["cargo", "bloat", "--version"], cwd=repo_root)
    has_cargo_bloat = cargo_bloat_probe.returncode == 0
    cargo_bloat_version: str | None = None
    if has_cargo_bloat:
        combined = (cargo_bloat_probe.stdout + ("\n" + cargo_bloat_probe.stderr if cargo_bloat_probe.stderr else "")).strip()
        if combined:
            cargo_bloat_version = combined.splitlines()[0].strip() or None

    crates_cmd: list[str] | None = None
    crates_out: CmdResult | None = None
    symbols_cmd: list[str] | None = None
    symbols_out: CmdResult | None = None

    tool_note: str | None = None
    if has_cargo_bloat:
        crates_cmd = [
            "cargo",
            "bloat",
            "-p",
            package,
            "--bin",
            bin_name,
            "--features",
            features,
            "--release",
            "--crates",
            "-n",
            str(args.top),
        ]
        symbols_cmd = [
            "cargo",
            "bloat",
            "-p",
            package,
            "--bin",
            bin_name,
            "--features",
            features,
            "--release",
            "-n",
            str(args.top),
        ]
        if target:
            crates_cmd.extend(["--target", target])
            symbols_cmd.extend(["--target", target])

        crates_out = _run_capture(crates_cmd, cwd=repo_root)
        symbols_out = _run_capture(symbols_cmd, cwd=repo_root)

        # If cargo-bloat failed for some reason, keep going to provide a fallback (llvm-size).
        if crates_out.returncode != 0 or symbols_out.returncode != 0:
            tool_note = (
                "Note: `cargo bloat` exited with a non-zero status. See the output below; "
                "falling back to `llvm-size` where available."
            )
    else:
        tool_note = (
            "Note: `cargo-bloat` is not installed (or `cargo bloat` is unavailable). "
            "Install with `cargo install cargo-bloat --locked` for crate/symbol breakdowns."
        )

    if over_limit:
        prefix = tool_note + "\n\n" if tool_note else ""
        tool_note = (
            prefix
            + f"Warning: release binary exceeds configured limit (**{_human_bytes(bin_size_bytes)}** > **{limit_mb} MB**)."
        )

    if stripped is False:
        prefix = tool_note + "\n\n" if tool_note else ""
        tool_note = (
            prefix
            + "Note: the release binary appears to be **not stripped**. "
            "Stripping debug symbols is often a quick win for installer/binary size.\n\n"
            "Consider:\n"
            "- enabling `strip = true` (or `strip = \"symbols\"`) in `[profile.release]` (Cargo.toml)\n"
            "- or running `strip` on the final artifact in packaging/build steps"
        )

    # Optional: add llvm-size output (helpful even when cargo-bloat is present).
    llvm_size_cmd: list[str] | None = None
    llvm_size_out: CmdResult | None = None
    llvm_size = shutil.which("llvm-size") or shutil.which("size")
    if llvm_size:
        # Prefer SysV output when supported (section-by-section breakdown).
        llvm_size_cmd = [llvm_size, "-A", str(bin_path)]
        llvm_size_out = _run_capture(llvm_size_cmd, cwd=repo_root)

        # Some `size` implementations don't support `-A`. Retry without it if needed.
        if llvm_size_out.returncode != 0:
            llvm_size_cmd = [llvm_size, str(bin_path)]
            llvm_size_out = _run_capture(llvm_size_cmd, cwd=repo_root)

    md = _render_markdown(
        package=package,
        bin_name=bin_name,
        features=features,
        target=target,
        repo_root=repo_root,
        target_dir=target_dir,
        bin_path=bin_path,
        bin_size_bytes=bin_size_bytes,
        limit_mb=limit_mb,
        enforce=enforce,
        rustc_version=rustc_version,
        cargo_version=cargo_version,
        git_sha=git_sha,
        git_ref=git_ref,
        cargo_bloat_version=cargo_bloat_version,
        file_info=file_info,
        stripped=stripped,
        build_ran=build_ran,
        json_out_path=json_out_path,
        build_cmd=build_cmd,
        crates_cmd=crates_cmd,
        crates_out=crates_out,
        symbols_cmd=symbols_cmd,
        symbols_out=symbols_out,
        llvm_size_cmd=llvm_size_cmd,
        llvm_size_out=llvm_size_out,
        tool_note=tool_note,
    )

    print(md)
    sys.stdout.flush()
    _append_step_summary(md)

    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(md, encoding="utf-8")
    if args.json_out:
        llvm_tool = Path(llvm_size_cmd[0]).name if llvm_size_cmd else None
        _write_json_out(
            args.json_out,
            {
                "status": "ok",
                "generated_at": _dt.datetime.now(tz=_dt.timezone.utc).isoformat(),
                "runner": {"os": os.environ.get("RUNNER_OS") or platform.system(), "arch": os.environ.get("RUNNER_ARCH") or platform.machine()},
                "git": {"sha": git_sha, "ref": git_ref},
                "package": package,
                "bin_name": bin_name,
                "features": features,
                "target": target,
                "target_dir": _relpath(target_dir, repo_root),
                "bin_path": _relpath(bin_path, repo_root),
                "bin_size_bytes": bin_size_bytes,
                "limit_mb": limit_mb,
                "enforce": enforce,
                "over_limit": over_limit,
                "toolchain": {"rustc": rustc_version, "cargo": cargo_version},
                "tools": {
                    "cargo_bloat": cargo_bloat_version,
                    "llvm_size": llvm_tool,
                    "file": file_info,
                    "stripped": stripped,
                },
                "commands": {
                    "build": build_cmd,
                    "cargo_bloat_crates": crates_cmd,
                    "cargo_bloat_symbols": symbols_cmd,
                    "llvm_size": llvm_size_cmd,
                },
                "outputs": {
                    "cargo_bloat_crates": crates_out.combined if crates_out else None,
                    "cargo_bloat_symbols": symbols_out.combined if symbols_out else None,
                    "llvm_size": llvm_size_out.combined if llvm_size_out else None,
                },
                "build_ran": build_ran,
            },
        )

    # Informational by default: only fail when explicitly enforcing a size limit.
    if enforce and over_limit:
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
