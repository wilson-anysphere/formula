#!/usr/bin/env python3
"""
Validate that packaged desktop bundles include open-source/compliance artifacts.

This script is intended to run *after* a Tauri release build has produced bundles under
`target/**/release/bundle`.

Checks (by platform):
- macOS: Mount the generated DMG(s) and assert `Formula.app/Contents/Resources/` contains
  `LICENSE` and `NOTICE`.
- Linux: Extract AppImage / DEB / RPM (if present) and assert `usr/share/doc/<package>/`
  contains `LICENSE` and `NOTICE`.
- Windows: Best-effort extraction of NSIS/MSI installers and assert they contain `LICENSE`
  and `NOTICE` somewhere in their payload.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence


REQUIRED_FILENAMES: tuple[str, ...] = ("LICENSE", "NOTICE")


class ValidationError(RuntimeError):
    pass


def _run(cmd: Sequence[str], *, cwd: Path | None = None) -> str:
    try:
        out = subprocess.check_output(
            list(cmd),
            cwd=str(cwd) if cwd is not None else None,
            stderr=subprocess.STDOUT,
            text=True,
        )
    except subprocess.CalledProcessError as exc:
        msg = (
            f"Command failed (exit {exc.returncode}): {' '.join(cmd)}\n"
            f"--- output ---\n{exc.output}\n--- end output ---"
        )
        raise ValidationError(msg) from exc
    return out


def _iter_bundle_files(bundle_dir: Path) -> Iterable[Path]:
    # Don't traverse huge app bundles / debug symbol directories.
    skip_suffixes = (".app", ".dSYM", ".framework", ".xcframework")
    for root, dirs, files in os.walk(bundle_dir):
        dirs[:] = [d for d in dirs if not d.endswith(skip_suffixes)]
        for f in files:
            yield Path(root) / f


def _candidate_target_dirs(repo_root: Path) -> list[Path]:
    candidates: list[Path] = []
    for p in (
        repo_root / "apps" / "desktop" / "src-tauri" / "target",
        repo_root / "target",
    ):
        if p.is_dir():
            candidates.append(p)
    return candidates


def _find_bundle_dirs(repo_root: Path) -> list[Path]:
    bundle_dirs: list[Path] = []
    for t in _candidate_target_dirs(repo_root):
        bundle_dirs.extend([p for p in t.glob("**/release/bundle") if p.is_dir()])

    # De-dupe while preserving order.
    seen: set[Path] = set()
    uniq: list[Path] = []
    for b in bundle_dirs:
        try:
            key = b.resolve()
        except FileNotFoundError:
            key = b
        if key in seen:
            continue
        seen.add(key)
        uniq.append(b)
    return uniq


@dataclass(frozen=True)
class BundleArtifacts:
    dmgs: list[Path]
    appimages: list[Path]
    debs: list[Path]
    rpms: list[Path]
    exes: list[Path]
    msis: list[Path]


def _gather_artifacts(bundle_dirs: list[Path]) -> BundleArtifacts:
    dmgs: list[Path] = []
    appimages: list[Path] = []
    debs: list[Path] = []
    rpms: list[Path] = []
    exes: list[Path] = []
    msis: list[Path] = []

    for bundle_dir in bundle_dirs:
        for path in _iter_bundle_files(bundle_dir):
            if not path.is_file():
                continue
            name = path.name
            lower = name.lower()
            if lower.endswith(".dmg"):
                dmgs.append(path)
            elif name.endswith(".AppImage"):
                appimages.append(path)
            elif lower.endswith(".deb"):
                debs.append(path)
            elif lower.endswith(".rpm"):
                rpms.append(path)
            elif lower.endswith(".exe"):
                exes.append(path)
            elif lower.endswith(".msi"):
                msis.append(path)

    # Sort for stable output.
    return BundleArtifacts(
        dmgs=sorted(dmgs),
        appimages=sorted(appimages),
        debs=sorted(debs),
        rpms=sorted(rpms),
        exes=sorted(exes),
        msis=sorted(msis),
    )


def _find_required_file(root: Path, filename: str) -> Path | None:
    direct = root / filename
    if direct.is_file():
        return direct
    for p in root.rglob(filename):
        if p.is_file():
            return p
    return None


def _find_7z() -> str | None:
    for name in ("7z", "7z.exe"):
        p = shutil.which(name)
        if p:
            return p
    candidates = [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ]
    for c in candidates:
        if Path(c).is_file():
            return c
    return None


def _tauri_main_binary_name(repo_root: Path) -> str:
    cfg_path = repo_root / "apps" / "desktop" / "src-tauri" / "tauri.conf.json"
    try:
        data = json.loads(cfg_path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ValidationError(f"Missing Tauri config: {cfg_path}") from exc
    val = data.get("mainBinaryName")
    if not isinstance(val, str) or not val.strip():
        raise ValidationError(f"Expected a string mainBinaryName in {cfg_path}")
    return val.strip()


def _validate_macos_dmg(dmg_path: Path) -> None:
    if sys.platform != "darwin":
        raise ValidationError("macOS DMG validation can only run on macOS runners.")

    if not dmg_path.is_file():
        raise ValidationError(f"DMG not found: {dmg_path}")

    with tempfile.TemporaryDirectory(prefix="formula-dmg-") as td:
        mountpoint = Path(td) / "mnt"
        mountpoint.mkdir(parents=True, exist_ok=True)
        _run(
            [
                "hdiutil",
                "attach",
                "-nobrowse",
                "-readonly",
                "-noverify",
                "-mountpoint",
                str(mountpoint),
                str(dmg_path),
            ]
        )
        try:
            apps = list(mountpoint.glob("*.app"))
            if not apps:
                apps = list(mountpoint.rglob("*.app"))
            if not apps:
                raise ValidationError(f"No .app found inside mounted DMG: {dmg_path}")

            # Usually only one app in the DMG. Validate all to be safe.
            for app in apps:
                resources_dir = app / "Contents" / "Resources"
                if not resources_dir.is_dir():
                    raise ValidationError(f"Missing Resources dir: {resources_dir} (from {dmg_path})")
                for filename in REQUIRED_FILENAMES:
                    found = _find_required_file(resources_dir, filename)
                    if not found:
                        raise ValidationError(
                            f"Missing {filename} in macOS app bundle Resources: {resources_dir} (from {dmg_path})"
                        )
        finally:
            # Detach the DMG even if validation fails.
            try:
                _run(["hdiutil", "detach", str(mountpoint)])
            except ValidationError:
                # Best-effort force detach (avoid masking the real validation error).
                subprocess.run(["hdiutil", "detach", "-force", str(mountpoint)], check=False)


def _validate_linux_deb(deb_path: Path, doc_dir_rel: Path) -> None:
    if not deb_path.is_file():
        raise ValidationError(f"DEB not found: {deb_path}")
    with tempfile.TemporaryDirectory(prefix="formula-deb-") as td:
        out_dir = Path(td) / "root"
        out_dir.mkdir(parents=True, exist_ok=True)
        _run(["dpkg-deb", "-x", str(deb_path), str(out_dir)])
        doc_dir = out_dir / doc_dir_rel
        for filename in REQUIRED_FILENAMES:
            p = doc_dir / filename
            if not p.is_file():
                raise ValidationError(f"Missing {filename} in extracted DEB at: {p} (from {deb_path})")


def _extract_rpm(rpm_path: Path, out_dir: Path) -> None:
    rpm2cpio = shutil.which("rpm2cpio")
    if rpm2cpio:
        # rpm2cpio writes a cpio archive to stdout.
        proc_rpm2cpio = subprocess.Popen(
            [rpm2cpio, str(rpm_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        assert proc_rpm2cpio.stdout is not None  # for type checkers
        assert proc_rpm2cpio.stderr is not None  # for type checkers
        try:
            cpio_proc = subprocess.run(
                ["cpio", "-idmv"],
                cwd=str(out_dir),
                stdin=proc_rpm2cpio.stdout,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
        finally:
            # Ensure rpm2cpio can receive SIGPIPE and exit once cpio is done.
            proc_rpm2cpio.stdout.close()

        rpm2cpio_stderr = proc_rpm2cpio.stderr.read().decode("utf-8", errors="replace")
        rpm2cpio_rc = proc_rpm2cpio.wait()

        if cpio_proc.returncode != 0 or rpm2cpio_rc != 0:
            raise ValidationError(
                f"Failed to extract RPM via rpm2cpio/cpio: {rpm_path}\n"
                f"cpio exit code: {cpio_proc.returncode}\n"
                f"rpm2cpio exit code: {rpm2cpio_rc}\n"
                f"cpio output:\n{cpio_proc.stdout}\n"
                f"rpm2cpio stderr:\n{rpm2cpio_stderr}"
            )
        return

    bsdtar = shutil.which("bsdtar") or shutil.which("tar")
    if bsdtar:
        # libarchive can usually unpack RPM directly.
        _run([bsdtar, "-xf", str(rpm_path), "-C", str(out_dir)])
        return

    raise ValidationError(
        "Cannot extract RPM: missing `rpm2cpio` (recommended) or `bsdtar`. "
        f"RPM artifact was: {rpm_path}"
    )


def _validate_linux_rpm(rpm_path: Path, doc_dir_rel: Path) -> None:
    if not rpm_path.is_file():
        raise ValidationError(f"RPM not found: {rpm_path}")
    with tempfile.TemporaryDirectory(prefix="formula-rpm-") as td:
        out_dir = Path(td) / "root"
        out_dir.mkdir(parents=True, exist_ok=True)
        _extract_rpm(rpm_path, out_dir)
        doc_dir = out_dir / doc_dir_rel
        for filename in REQUIRED_FILENAMES:
            p = doc_dir / filename
            if not p.is_file():
                raise ValidationError(f"Missing {filename} in extracted RPM at: {p} (from {rpm_path})")


def _validate_linux_appimage(appimage_path: Path, doc_dir_rel: Path) -> None:
    if not appimage_path.is_file():
        raise ValidationError(f"AppImage not found: {appimage_path}")

    # AppImage extract requires the file to be executable.
    try:
        appimage_path.chmod(appimage_path.stat().st_mode | 0o111)
    except PermissionError:
        pass

    with tempfile.TemporaryDirectory(prefix="formula-appimage-") as td:
        td_path = Path(td)
        _run([str(appimage_path), "--appimage-extract"], cwd=td_path)
        root = td_path / "squashfs-root"
        if not root.is_dir():
            raise ValidationError(f"Expected squashfs-root after AppImage extract: {appimage_path}")
        doc_dir = root / doc_dir_rel
        for filename in REQUIRED_FILENAMES:
            p = doc_dir / filename
            if not p.is_file():
                raise ValidationError(f"Missing {filename} in extracted AppImage at: {p} (from {appimage_path})")


def _validate_windows_installer_payload(installer_path: Path) -> None:
    seven_zip = _find_7z()
    if not seven_zip:
        raise ValidationError(
            "7-Zip (7z) not found; cannot validate Windows installer payload. "
            f"Installer was: {installer_path}"
        )

    with tempfile.TemporaryDirectory(prefix="formula-win-installer-") as td:
        out_dir = Path(td) / "extract"
        out_dir.mkdir(parents=True, exist_ok=True)

        # Best-effort extraction:
        # - For MSI, prefer `msiexec /a` to get the installed file layout.
        # - Otherwise, use 7z to unpack.
        if installer_path.name.lower().endswith(".msi"):
            targetdir = out_dir / "msi-admin"
            targetdir.mkdir(parents=True, exist_ok=True)
            target = str(targetdir)
            if not target.endswith(("\\", "/")):
                target += "\\"
            try:
                _run(["msiexec", "/a", str(installer_path), "/qn", f"TARGETDIR={target}"])
                search_root = targetdir
            except ValidationError:
                _run([seven_zip, "x", "-y", f"-o{out_dir}", str(installer_path)])
                search_root = out_dir
        else:
            _run([seven_zip, "x", "-y", f"-o{out_dir}", str(installer_path)])
            search_root = out_dir

        for filename in REQUIRED_FILENAMES:
            found = _find_required_file(search_root, filename)
            if not found:
                raise ValidationError(
                    f"Missing {filename} in extracted Windows installer payload: {installer_path}\n"
                    f"Searched under: {search_root}"
                )


def _validate_current_platform(artifacts: BundleArtifacts, repo_root: Path) -> None:
    if sys.platform == "darwin":
        if not artifacts.dmgs:
            raise ValidationError(
                "No DMG artifacts found to validate.\n"
                "Hint: run a macOS Tauri build first (tauri build should produce target/release/bundle/dmg/*.dmg)."
            )
        for dmg in artifacts.dmgs:
            print(f"bundle-validate: macOS DMG: {dmg}")
            _validate_macos_dmg(dmg)
        return

    if sys.platform.startswith("linux"):
        main_binary = _tauri_main_binary_name(repo_root)
        doc_dir_rel = Path("usr") / "share" / "doc" / main_binary

        if not (artifacts.appimages or artifacts.debs or artifacts.rpms):
            raise ValidationError(
                "No Linux bundle artifacts found to validate (.AppImage/.deb/.rpm).\n"
                "Hint: run a Linux Tauri build first (tauri build should produce target/release/bundle/*)."
            )
        for appimage in artifacts.appimages:
            print(f"bundle-validate: Linux AppImage: {appimage}")
            _validate_linux_appimage(appimage, doc_dir_rel)
        for deb in artifacts.debs:
            print(f"bundle-validate: Linux DEB: {deb}")
            _validate_linux_deb(deb, doc_dir_rel)
        for rpm in artifacts.rpms:
            print(f"bundle-validate: Linux RPM: {rpm}")
            _validate_linux_rpm(rpm, doc_dir_rel)
        return

    if sys.platform == "win32":
        installers = artifacts.exes + artifacts.msis
        if not installers:
            raise ValidationError(
                "No Windows installer artifacts found to validate (.exe/.msi).\n"
                "Hint: run a Windows Tauri build first (tauri build should produce target/release/bundle/*)."
            )
        for inst in installers:
            print(f"bundle-validate: Windows installer: {inst}")
            _validate_windows_installer_payload(inst)
        return

    raise ValidationError(f"Unsupported platform for desktop bundle validation: {sys.platform}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate desktop release bundles include LICENSE + NOTICE artifacts."
    )
    parser.add_argument(
        "--bundle-dir",
        type=Path,
        action="append",
        default=None,
        help="Explicit bundle directory to scan (can be specified multiple times). "
        "If omitted, auto-detects target/**/release/bundle under common locations.",
    )
    args = parser.parse_args()

    repo_root = Path.cwd()

    bundle_dirs: list[Path] = []
    if args.bundle_dir:
        bundle_dirs = [d for d in args.bundle_dir if d.is_dir()]
        missing = [str(d) for d in args.bundle_dir if not d.is_dir()]
        if missing:
            print(f"bundle-validate: ERROR --bundle-dir not found: {', '.join(missing)}", file=sys.stderr)
            return 2
    else:
        bundle_dirs = _find_bundle_dirs(repo_root)

    if not bundle_dirs:
        expected = repo_root / "apps" / "desktop" / "src-tauri" / "target" / "release" / "bundle"
        print(
            "bundle-validate: ERROR No Tauri bundle directories found.\n"
            f"Searched for: target/**/release/bundle (from {repo_root})\n"
            f"Expected a directory like: {expected}\n"
            "Hint: build the desktop app with `(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)` "
            "before running this script.",
            file=sys.stderr,
        )
        return 1

    artifacts = _gather_artifacts(bundle_dirs)
    try:
        _validate_current_platform(artifacts, repo_root=repo_root)
    except ValidationError as exc:
        print(f"bundle-validate: ERROR {exc}", file=sys.stderr)
        return 1

    print("bundle-validate: OK (LICENSE/NOTICE present in bundles)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
