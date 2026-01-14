#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path


def repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def extract_pyodide_version(ensure_script: Path) -> str:
    src = ensure_script.read_text(encoding="utf-8")
    m = re.search(r"const\s+PYODIDE_VERSION\s*=\s*['\"]([^'\"]+)['\"]", src)
    if not m:
        raise SystemExit(
            f"pyodide-cache check: failed to find PYODIDE_VERSION in {ensure_script.as_posix()}"
        )
    return m.group(1)


def parse_jobs(yml_path: Path) -> dict[str, list[tuple[int, str]]]:
    """Best-effort workflow job block parser (no YAML dependencies).

    This is intentionally lightweight: we just split the `jobs:` mapping into job id
    blocks using indentation heuristics (2-space job keys).
    """

    jobs: dict[str, list[tuple[int, str]]] = {}
    lines = yml_path.read_text(encoding="utf-8").splitlines()

    in_jobs = False
    current_job: str | None = None

    for idx, line in enumerate(lines, start=1):
        if not in_jobs:
            if line.strip() == "jobs:":
                in_jobs = True
            continue

        # Exit the jobs section when we hit a new top-level key.
        if line and not line.startswith(" ") and not line.startswith("#"):
            break

        m = re.match(r"^  ([A-Za-z0-9_-]+):\s*(#.*)?$", line)
        if m:
            current_job = m.group(1)
            jobs.setdefault(current_job, [])
            continue

        if current_job is not None:
            jobs[current_job].append((idx, line))

    return jobs


def check_job(
    workflow: Path, job_id: str, job_lines: list[tuple[int, str]], pyodide_version: str
) -> list[str]:
    errors: list[str] = []
    text = "\n".join(line for _, line in job_lines)

    def is_comment(line: str) -> bool:
        stripped = line.lstrip()
        return stripped.startswith("#")

    def is_desktop_build_cmd(line: str) -> bool:
        # Match both:
        # - `run: pnpm build:desktop` (desktop frontend build)
        # - `cargo tauri build` / `tauri build` (Tauri bundles invoke `pnpm build` via beforeBuildCommand)
        if is_comment(line):
            return False
        return bool(
            re.search(r"\bpnpm\b.*\bbuild:desktop\b", line)
            or re.search(r"\btauri\s+build\b", line)
        )

    build_lines = [ln for ln, line in job_lines if is_desktop_build_cmd(line)]
    if not build_lines:
        return errors

    def require(substr: str, desc: str) -> None:
        if substr not in text:
            errors.append(f"- Missing {desc} ({substr})")

    require("Restore Pyodide asset cache", "Pyodide cache restore step")
    require("Ensure Pyodide assets are present", "Pyodide ensure/download step")
    require("Save Pyodide asset cache", "Pyodide cache save step")

    # Ensure ordering: restore must happen before build:desktop.
    restore_line = next(
        (ln for ln, line in job_lines if "Restore Pyodide asset cache" in line), None
    )
    if restore_line is not None:
        for build_ln in build_lines:
            if restore_line > build_ln:
                errors.append(
                    f"- Pyodide cache restore step appears after the desktop build command (restore line {restore_line}, build line {build_ln})"
                )

    # Ensure the cache key includes OS + version + ensure script hash.
    if "key: pyodide-" not in text:
        errors.append("- Missing `key: pyodide-...` in cache configuration")
    else:
        if "runner.os" not in text:
            errors.append("- Pyodide cache key is missing runner.os")
        if "hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs')" not in text:
            errors.append("- Pyodide cache key is missing hashFiles(ensure-pyodide-assets.mjs)")
        if "steps.pyodide.outputs.version" not in text and pyodide_version not in text:
            errors.append(
                f"- Pyodide cache key/path is missing version reference (expected steps.pyodide.outputs.version or {pyodide_version})"
            )

    # Ensure we cache the expected directory.
    if "apps/desktop/public/pyodide/" not in text:
        errors.append("- Pyodide cache config is missing apps/desktop/public/pyodide/ path")

    if errors:
        header = f"{workflow.as_posix()} job `{job_id}` runs a desktop build but is missing required Pyodide caching:"
        return [header, *errors]

    return []


def main() -> None:
    root = repo_root()
    pyodide_version = extract_pyodide_version(
        root / "apps" / "desktop" / "scripts" / "ensure-pyodide-assets.mjs"
    )

    workflows_dir = root / ".github" / "workflows"
    failures: list[str] = []
    for wf in sorted(workflows_dir.glob("*.yml")):
        jobs = parse_jobs(wf)
        for job_id, job_lines in jobs.items():
            failures.extend(check_job(wf, job_id, job_lines, pyodide_version))

    if failures:
        msg = "\n".join(failures)
        raise SystemExit(msg)

    print("Pyodide cache guard: OK")


if __name__ == "__main__":
    main()
