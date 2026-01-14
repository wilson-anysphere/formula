#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path


def repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def strip_wrapping_quotes(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value


def strip_flow_mapping_trailing_brace(value: str) -> str:
    """Remove a trailing `}` introduced by inline YAML flow mappings.

    Some workflows use:

        with: { path: ..., key: ... }

    When we extract the `key:`/`path:` value by splitting on commas, the closing
    `}` can end up attached to the final value. We want to drop *one* mapping
    brace without clobbering GitHub expression braces (e.g. `${{ ... }}`).
    """

    value = value.strip()
    if value.endswith("}}}"):
        return value[:-1].rstrip()
    if value.endswith("}") and not value.endswith("}}"):
        return value[:-1].rstrip()
    return value


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
        # - `run: pnpm build:desktop` (workspace desktop frontend build)
        # - `run: pnpm -C apps/desktop build` (package-local desktop frontend build)
        # - `cargo tauri build` / `tauri build` (Tauri bundles invoke `pnpm build` via beforeBuildCommand)
        # - `tauri-apps/tauri-action` (runs a Tauri bundle build, which triggers beforeBuildCommand)
        if is_comment(line):
            return False
        return bool(
            re.search(r"\bpnpm\b.*\bbuild:desktop\b", line)
            or re.search(r"\bpnpm\b.*\b(?:-C|--dir)\b\s+apps/desktop\b.*\bbuild(?=$|\s)", line)
            or re.search(r"\btauri\s+build\b", line)
            or re.search(r"\btauri-apps/tauri-action\b", line)
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
    pyodide_key_lines: list[tuple[int, str]] = []
    for ln, line in job_lines:
        if is_comment(line):
            continue
        # Strip trailing inline comments.
        line = line.split("#", 1)[0]
        for m in re.finditer(r"\bkey:\s*([^,]+)", line):
            raw_value = m.group(1).strip()
            raw_value = strip_flow_mapping_trailing_brace(raw_value)
            raw_value = strip_wrapping_quotes(raw_value)
            if raw_value.startswith("pyodide"):
                pyodide_key_lines.append((ln, raw_value))

    if not pyodide_key_lines:
        errors.append("- Missing `key: pyodide-...` in cache configuration")
    else:
        for ln, value in pyodide_key_lines:
            if "runner.os" not in value:
                errors.append(f"- Pyodide cache key (line {ln}) is missing runner.os")
            if "hashFiles('apps/desktop/scripts/ensure-pyodide-assets.mjs')" not in value:
                errors.append(
                    f"- Pyodide cache key (line {ln}) is missing hashFiles(ensure-pyodide-assets.mjs)"
                )
            if "steps.pyodide.outputs.version" not in value and pyodide_version not in value:
                errors.append(
                    f"- Pyodide cache key (line {ln}) is missing version reference (expected steps.pyodide.outputs.version or {pyodide_version})"
                )

    # Ensure we cache the expected directory.
    #
    # Most workflows cache the versioned `full/` directory that `ensure-pyodide-assets.mjs`
    # downloads into. Some release/dry-run workflows cache the entire
    # `apps/desktop/public/pyodide/` tree to allow restoring tracked files + cross-OS
    # cache sharing.
    expected_dirs = [
        "apps/desktop/public/pyodide/v${{ steps.pyodide.outputs.version }}/full",
        f"apps/desktop/public/pyodide/v{pyodide_version}/full",
        "apps/desktop/public/pyodide/",
        "apps/desktop/public/pyodide",
    ]
    if not any(d in text for d in expected_dirs):
        errors.append(
            "- Pyodide cache config is missing an expected pyodide directory path "
            "(apps/desktop/public/pyodide/ or apps/desktop/public/pyodide/v${{ steps.pyodide.outputs.version }}/full/)"
        )

    # If we cache the entire `apps/desktop/public/pyodide/` tree (instead of just the
    # versioned `full/` directory), the cache may include tracked repo files like
    # `README.md` and the versioned `.gitignore`. Restoring an older cache could
    # overwrite those tracked files and trigger reproducibility guards (`git diff
    # --exit-code`) or cause confusing dirty-worktree states.
    #
    # Require a step that restores the tracked files back to HEAD after the cache is
    # restored while keeping the untracked downloaded assets intact.

    def split_steps() -> list[list[tuple[int, str]]]:
        """Split a job block into step blocks using indentation heuristics.

        We only need a best-effort parser for guardrails (not full YAML). This
        splits on lines like:

        - `- name: ...`
        - `- uses: ...`

        at a consistent indentation level.
        """

        steps: list[list[tuple[int, str]]] = []
        i = 0
        while i < len(job_lines):
            ln, line = job_lines[i]
            if is_comment(line):
                i += 1
                continue
            m = re.match(r"^(\s*)-\s+(name|uses):\s*(.+)$", line)
            if not m:
                i += 1
                continue
            indent = len(m.group(1))
            start = i
            i += 1
            while i < len(job_lines):
                _, next_line = job_lines[i]
                if is_comment(next_line):
                    i += 1
                    continue
                m2 = re.match(r"^(\s*)-\s+(name|uses):\s*(.+)$", next_line)
                if m2 and len(m2.group(1)) == indent:
                    break
                i += 1
            steps.append(job_lines[start:i])
        return steps

    def find_step(name_substr: str) -> list[tuple[int, str]] | None:
        for step in split_steps():
            for _, line in step:
                if is_comment(line):
                    continue
                if "name:" in line and name_substr in line:
                    return step
        return None

    def extract_multiline_value(
        step: list[tuple[int, str]], field: str
    ) -> tuple[int, list[str]] | None:
        """Extract `field:` value lines from a `|`-style multi-line scalar.

        Returns (line_number_of_field, values) or None if field not present.
        """

        # YAML block scalar indicators:
        # - `|` / `>` (literal/folded)
        # - optional indentation + chomping indicators in either order (e.g. `|2-`, `|-2`).
        # GitHub Actions workflow YAML commonly uses these forms, so treat anything of the form
        # `[|>][0-9+-]*` as a block scalar header.
        block_marker_re = re.compile(r"^[|>][0-9+-]*$")
        for i, (ln, line) in enumerate(step):
            if is_comment(line):
                continue
            m = re.match(rf"^(\s*){re.escape(field)}:\s*(.*)$", line)
            if not m:
                continue
            indent = len(m.group(1))
            rest = m.group(2).strip()
            if rest and not block_marker_re.match(rest):
                # Single-line value.
                rest = strip_flow_mapping_trailing_brace(strip_wrapping_quotes(rest))
                return (ln, [rest])
            values: list[str] = []
            for _, next_line in step[i + 1 :]:
                if next_line.strip() == "" or is_comment(next_line):
                    continue
                next_indent = len(next_line) - len(next_line.lstrip(" "))
                if next_indent <= indent:
                    # Dedented; stop.
                    break
                value = next_line.strip()
                # Allow YAML list forms (`field:\n  - foo\n  - bar`).
                if value.startswith("- "):
                    value = value[2:].strip()
                # Strip inline comments in value lines (best-effort).
                value = value.split("#", 1)[0].strip()
                value = strip_flow_mapping_trailing_brace(strip_wrapping_quotes(value))
                if value:
                    values.append(value)
            return (ln, values)
        return None

    def extract_step_inline_field(step: list[tuple[int, str]], field: str) -> list[str]:
        """Extract field values from inline `with: { ... }` flow mappings (best-effort)."""

        out: list[str] = []
        # Match `field: <value>` inside an inline mapping. Capture until the next comma
        # or end-of-line; we'll strip the trailing `}` below if present.
        pattern = re.compile(rf"\b{re.escape(field)}:\s*([^,]+)")
        for _, line in step:
            if is_comment(line):
                continue
            candidate = line.split("#", 1)[0]
            for m in pattern.finditer(candidate):
                value = m.group(1).strip()
                value = strip_flow_mapping_trailing_brace(strip_wrapping_quotes(value))
                if value:
                    out.append(value)
        return out

    def step_caches_whole_pyodide_tree(step: list[tuple[int, str]] | None) -> bool:
        if step is None:
            return False
        extracted = extract_multiline_value(step, "path")
        paths: list[str] = []
        if extracted is not None:
            _, values = extracted
            paths.extend(values)
        paths.extend(extract_step_inline_field(step, "path"))
        for raw in paths:
            cleaned = strip_flow_mapping_trailing_brace(strip_wrapping_quotes(raw)).strip()
            cleaned = cleaned.rstrip("/")
            if cleaned == "apps/desktop/public/pyodide":
                return True
        return False

    restore_step = find_step("Restore Pyodide asset cache")
    save_step = find_step("Save Pyodide asset cache")
    caches_whole_pyodide_tree = step_caches_whole_pyodide_tree(
        restore_step
    ) or step_caches_whole_pyodide_tree(save_step)

    if caches_whole_pyodide_tree:
        restore_tracked_line = next(
            (
                ln
                for ln, line in job_lines
                if not is_comment(line)
                and "git restore --source=HEAD -- apps/desktop/public/pyodide" in line
            ),
            None,
        )
        if restore_tracked_line is None:
            errors.append(
                "- Job caches apps/desktop/public/pyodide/ but does not restore tracked files after cache restore "
                "(expected a `git restore --source=HEAD -- apps/desktop/public/pyodide/` step)"
            )
        elif restore_line is not None and restore_tracked_line < restore_line:
            errors.append(
                f"- Tracked Pyodide restore step appears before the cache restore step (tracked restore line {restore_tracked_line}, cache restore line {restore_line})"
            )

        # Cross-OS restore requires both:
        # - `restore-keys` that can match a cache created on a different OS
        # - `enableCrossOsArchive: true` on restore/save
        if restore_step is None:
            errors.append("- Missing 'Restore Pyodide asset cache' step block (unable to validate enableCrossOsArchive/restore-keys)")
        else:
            if "enableCrossOsArchive: true" not in "\n".join(line for _, line in restore_step):
                errors.append(
                    "- Pyodide cache restore step is missing `enableCrossOsArchive: true` (required for cross-OS restore keys)"
                )
            restore_keys = extract_multiline_value(restore_step, "restore-keys")
            if restore_keys is None:
                errors.append(
                    "- Pyodide cache restore step is missing `restore-keys:` (required to fall back to same Pyodide version across OSes)"
                )
            else:
                _, keys = restore_keys
                has_cross_os_fallback = any(
                    ("runner.os" not in k)
                    and (
                        "steps.pyodide.outputs.version" in k
                        or pyodide_version in k
                    )
                    for k in keys
                )
                if not has_cross_os_fallback:
                    errors.append(
                        "- Pyodide cache restore-keys is missing a cross-OS fallback prefix (expected a key that includes the Pyodide version but not runner.os)"
                    )

        if save_step is None:
            errors.append("- Missing 'Save Pyodide asset cache' step block (unable to validate enableCrossOsArchive)")
        else:
            if "enableCrossOsArchive: true" not in "\n".join(line for _, line in save_step):
                errors.append(
                    "- Pyodide cache save step is missing `enableCrossOsArchive: true` (ensures caches are restorable across OSes)"
                )

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
