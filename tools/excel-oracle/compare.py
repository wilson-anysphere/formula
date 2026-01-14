#!/usr/bin/env python3
"""
Compare formula engine results against an Excel oracle dataset.

Input files:
- cases.json: the canonical case corpus (for looking up formula + inputs)
- expected.json: results from tools/excel-oracle/run-excel-oracle.ps1
- actual.json: results from the formula engine under test (same schema as expected)

Output:
- report JSON with mismatches and a summary.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
import urllib.parse
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


_PRIVACY_PUBLIC = "public"
_PRIVACY_PRIVATE = "private"

_KNOWN_FUNCTION_NAMES: set[str] | None = None


def _load_known_function_names() -> set[str]:
    global _KNOWN_FUNCTION_NAMES
    if _KNOWN_FUNCTION_NAMES is not None:
        return _KNOWN_FUNCTION_NAMES

    names: set[str] = set()
    try:
        repo_root = Path(__file__).resolve().parents[2]
        catalog_path = repo_root / "shared" / "functionCatalog.json"
        data = json.loads(catalog_path.read_text(encoding="utf-8"))
        funcs = data.get("functions") if isinstance(data, dict) else None
        if isinstance(funcs, list):
            for entry in funcs:
                if not isinstance(entry, dict):
                    continue
                name = entry.get("name")
                if isinstance(name, str) and name:
                    names.add(name.upper())
    except Exception:  # noqa: BLE001 (best-effort; privacy mode still works without allowlist)
        names = set()

    _KNOWN_FUNCTION_NAMES = names
    return names


def _redact_function_name(name: str, *, privacy_mode: str) -> str:
    if privacy_mode != _PRIVACY_PRIVATE:
        return name
    if not name or name.startswith("sha256="):
        return name
    known = _load_known_function_names()
    if name.upper() in known:
        return name
    return f"sha256={_sha256_text(name)}"


def _redact_tag_name(tag: str, *, privacy_mode: str) -> str:
    """Redact tag names that may embed custom function namespaces.

    Tags in the oracle corpus are typically either:
    - function names (e.g. "SUM", "XLOOKUP")
    - high-level categories (e.g. "arith", "spill")
    - or (in private corpora) custom add-in namespaces (e.g. "CORP.ADDIN.FOO")

    In privacy mode we keep category tags and built-in Excel function tags readable, but hash
    namespace-like tags (containing ".") that are not in the function allowlist.
    """

    if privacy_mode != _PRIVACY_PRIVATE:
        return tag
    if not tag or tag.startswith("sha256="):
        return tag

    raw = tag.strip()
    if "." not in raw:
        return tag

    normalized = raw.upper()
    if normalized.startswith("_XLFN."):
        normalized = normalized[len("_XLFN.") :]
    if normalized in _load_known_function_names():
        return tag
    return f"sha256={_sha256_text(normalized)}"


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _redact_text(value: str | None, *, privacy_mode: str) -> str | None:
    """Redact potentially sensitive free-form strings in privacy mode.

    This is used to scrub local filesystem paths embedded in reports (for example, Windows paths that
    include usernames like `C:\\Users\\Alice\\...`).
    """

    if not value or privacy_mode != _PRIVACY_PRIVATE:
        return value
    if value.startswith("sha256="):
        return value

    # Keep repo-relative paths readable; only hash strings that look like absolute filesystem paths
    # or URI-like paths (file://, smb://, etc).
    parsed = urllib.parse.urlparse(value)
    looks_abs = bool(
        value.startswith(("/", "\\", "~"))
        or value.startswith("//")
        or re.match(r"^[A-Za-z]:[\\/]", value)
        # urlparse treats `C:\foo` as scheme "c" on non-Windows too; that's fine (it is a path).
        or (parsed.scheme and ":" in value)
    )
    if not looks_abs:
        return value

    return f"sha256={_sha256_text(value)}"


def _redact_paths_in_obj(obj: Any, *, privacy_mode: str) -> Any:
    """Recursively redact values under keys named `path` or `*Path`."""

    if privacy_mode != _PRIVACY_PRIVATE:
        return obj
    if isinstance(obj, dict):
        out: dict[str, Any] = {}
        for k, v in obj.items():
            if (
                isinstance(k, str)
                and isinstance(v, str)
                and v
                and (k == "path" or k.endswith("Path"))
            ):
                out[k] = _redact_text(v, privacy_mode=privacy_mode)
            else:
                out[k] = _redact_paths_in_obj(v, privacy_mode=privacy_mode)
        return out
    if isinstance(obj, list):
        return [_redact_paths_in_obj(v, privacy_mode=privacy_mode) for v in obj]
    return obj


def _redact_error_details_in_obj(obj: Any, *, privacy_mode: str) -> Any:
    """Recursively redact free-form error `detail` strings in privacy mode.

    The formula engine (and some adapters) may include an additional `detail` field for error values
    (t="e") containing arbitrary text (often derived from exception strings). In privacy mode we
    hash these to avoid leaking filenames, local paths, or internal identifiers in CI artifacts.
    """

    if privacy_mode != _PRIVACY_PRIVATE:
        return obj
    if isinstance(obj, dict):
        out: dict[str, Any] = {}
        is_error_value = obj.get("t") == "e"
        for k, v in obj.items():
            if is_error_value and k == "detail" and isinstance(v, str) and v:
                out[k] = v if v.startswith("sha256=") else f"sha256={_sha256_text(v)}"
            else:
                out[k] = _redact_error_details_in_obj(v, privacy_mode=privacy_mode)
        return out
    if isinstance(obj, list):
        return [_redact_error_details_in_obj(v, privacy_mode=privacy_mode) for v in obj]
    return obj


def _redact_string_values_in_obj(obj: Any, *, privacy_mode: str) -> Any:
    """Recursively redact string values (t="s") that look like absolute paths/URIs.

    Excel-oracle mismatches include the expected/actual result values. When a case result is a
    string, it may contain a local filesystem path or a URI (for example, workbook content that
    embeds a hyperlink). In privacy mode, hash path/URI-like strings to avoid leaking local
    usernames/mount points or corporate domains in CI artifacts.
    """

    if privacy_mode != _PRIVACY_PRIVATE:
        return obj
    if isinstance(obj, dict):
        out: dict[str, Any] = {}
        is_string_value = obj.get("t") == "s"
        for k, v in obj.items():
            if is_string_value and k == "v" and isinstance(v, str) and v:
                out[k] = _redact_text(v, privacy_mode=privacy_mode) or v
            else:
                out[k] = _redact_string_values_in_obj(v, privacy_mode=privacy_mode)
        return out
    if isinstance(obj, list):
        return [_redact_string_values_in_obj(v, privacy_mode=privacy_mode) for v in obj]
    return obj


def _index_results(
    results: Iterable[dict[str, Any]], *, label: str
) -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    duplicates: set[str] = set()
    for r in results:
        cid = r.get("caseId")
        if not isinstance(cid, str):
            continue
        if cid in out:
            duplicates.add(cid)
        out[cid] = r
    if duplicates:
        preview = ", ".join(sorted(list(duplicates))[:25])
        suffix = "" if len(duplicates) <= 25 else f" (+{len(duplicates) - 25} more)"
        raise SystemExit(
            f"{label} dataset contains duplicate caseId entries ({len(duplicates)}): {preview}{suffix}"
        )
    return out


def _pretty_input(cell_input: dict[str, Any], *, privacy_mode: str) -> dict[str, Any]:
    if "formula" in cell_input:
        return {
            "cell": cell_input.get("cell"),
            "formula": _redact_formula(cell_input.get("formula"), privacy_mode=privacy_mode),
        }
    value = cell_input.get("value")
    if isinstance(value, str):
        value = _redact_text(value, privacy_mode=privacy_mode) or value
    return {"cell": cell_input.get("cell"), "value": value}


def _maybe_nonempty_str(value: Any) -> str | None:
    if isinstance(value, str) and value:
        return value
    return None


_FUNC_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")


def _extract_function_names(formula: str | None) -> list[str]:
    if not formula:
        return []
    raw = formula.strip()
    if raw.startswith("="):
        raw = raw[1:]

    out: list[str] = []
    for match in _FUNC_RE.finditer(raw):
        name = match.group(1).upper()
        if name.startswith("_XLFN."):
            name = name[len("_XLFN.") :]
        out.append(name)
    return out


def _redact_formula(formula: str | None, *, privacy_mode: str) -> str | None:
    """Redact non-standard/UDF function names inside a formula string.

    This is best-effort and intentionally conservative: it only rewrites function-call identifiers
    matched by `_FUNC_RE` and leaves the rest of the formula intact.
    """

    if not formula or privacy_mode != _PRIVACY_PRIVATE:
        return formula

    known = _load_known_function_names()

    def _replace(match: re.Match[str]) -> str:
        raw_name = match.group(1)
        suffix = match.group(0)[len(raw_name) :]  # includes any whitespace and the opening "("
        normalized = raw_name.upper()
        if normalized.startswith("_XLFN."):
            normalized = normalized[len("_XLFN.") :]
        if normalized in known:
            return raw_name + suffix
        return f"sha256={_sha256_text(normalized)}{suffix}"

    return _FUNC_RE.sub(_replace, formula)


@dataclass(frozen=True)
class CompareConfig:
    abs_tol: float
    rel_tol: float


def _parse_tag_tolerances(values: list[str], *, flag_name: str) -> dict[str, float]:
    """
    Parse `TAG=FLOAT` pairs into a mapping, taking the maximum for duplicate tags.
    """

    out: dict[str, float] = {}
    for raw in values:
        if not isinstance(raw, str) or "=" not in raw:
            raise SystemExit(
                f"Invalid {flag_name} value {raw!r}. Expected TAG=FLOAT (example: odd_coupon=1e-6)."
            )
        tag, value_str = raw.split("=", 1)
        tag = tag.strip()
        if not tag:
            raise SystemExit(
                f"Invalid {flag_name} value {raw!r}. Tag must be non-empty (example: odd_coupon=1e-6)."
            )
        try:
            value = float(value_str)
        except ValueError:
            raise SystemExit(
                f"Invalid {flag_name} value {raw!r}. {value_str!r} is not a float (example: odd_coupon=1e-6)."
            ) from None
        if not math.isfinite(value) or value < 0.0:
            raise SystemExit(
                f"Invalid {flag_name} value {raw!r}. Tolerance must be a finite, non-negative float."
            )

        prev = out.get(tag)
        if prev is None or value > prev:
            out[tag] = value
    return out


def _effective_cfg_for_tags(
    default: CompareConfig,
    *,
    tags: set[str],
    tag_abs_tol: dict[str, float],
    tag_rel_tol: dict[str, float],
) -> CompareConfig:
    abs_tol = default.abs_tol
    rel_tol = default.rel_tol
    for t in tags:
        v = tag_abs_tol.get(t)
        if v is not None and v > abs_tol:
            abs_tol = v
        v = tag_rel_tol.get(t)
        if v is not None and v > rel_tol:
            rel_tol = v
    return CompareConfig(abs_tol=abs_tol, rel_tol=rel_tol)


def _is_number(value_obj: Any) -> bool:
    return (
        isinstance(value_obj, dict)
        and value_obj.get("t") == "n"
        and isinstance(value_obj.get("v"), (int, float))
    )


def _numbers_close(a: float, b: float, cfg: CompareConfig) -> bool:
    # Handle NaN explicitly (Excel doesn't produce NaN, but engines might).
    if math.isnan(a) or math.isnan(b):
        return False
    return math.isclose(a, b, rel_tol=cfg.rel_tol, abs_tol=cfg.abs_tol)


def _compare_value(expected: Any, actual: Any, cfg: CompareConfig) -> tuple[bool, str, Any | None]:
    if expected == actual:
        return True, "ok", None

    if not isinstance(expected, dict) or not isinstance(actual, dict):
        return False, "type-mismatch", None

    et = expected.get("t")
    at = actual.get("t")
    if et != at:
        return False, "type-mismatch", None

    if et == "n":
        av = float(actual.get("v"))
        ev = float(expected.get("v"))
        return (_numbers_close(ev, av, cfg), "number-mismatch", None)

    if et in ("s", "b", "e"):
        return (expected.get("v") == actual.get("v"), f"{et}-mismatch", None)

    if et == "blank":
        return True, "ok", None

    if et == "arr":
        erows = expected.get("rows")
        arows = actual.get("rows")
        if not isinstance(erows, list) or not isinstance(arows, list):
            return False, "array-shape-mismatch", None
        if len(erows) != len(arows):
            return (
                False,
                "array-shape-mismatch",
                {"expectedRows": len(erows), "actualRows": len(arows)},
            )
        for r in range(len(erows)):
            if not isinstance(erows[r], list) or not isinstance(arows[r], list):
                return False, "array-shape-mismatch", {"row": r}
            if len(erows[r]) != len(arows[r]):
                return (
                    False,
                    "array-shape-mismatch",
                    {"row": r, "expectedCols": len(erows[r]), "actualCols": len(arows[r])},
                )
            for c in range(len(erows[r])):
                expected_cell = erows[r][c]
                actual_cell = arows[r][c]
                ok, reason, detail = _compare_value(expected_cell, actual_cell, cfg)
                if not ok:
                    return (
                        False,
                        f"array-mismatch:{reason}",
                        {
                            "row": r,
                            "col": c,
                            "reason": reason,
                            "detail": detail,
                            "expected": expected_cell,
                            "actual": actual_cell,
                        },
                    )
        return True, "ok", None

    return False, "unknown-type", None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cases", required=True, help="Path to cases.json")
    parser.add_argument("--expected", required=True, help="Path to Excel oracle results JSON")
    parser.add_argument("--actual", required=True, help="Path to engine results JSON")
    parser.add_argument("--report", required=True, help="Path to write mismatch report JSON")
    parser.add_argument(
        "--privacy-mode",
        choices=[_PRIVACY_PUBLIC, _PRIVACY_PRIVATE],
        default=_PRIVACY_PUBLIC,
        help=(
            "Control redaction of outputs. `private` hashes filesystem path metadata in the report "
            "(for example, paths that include local usernames/mount points) and hashes any non-standard "
            "formula function names surfaced in summary aggregations (e.g. topMissingFunctions)."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print how many cases would be compared (after tag filtering / max-cases) and exit without writing a report.",
    )
    parser.add_argument(
        "--include-tag",
        action="append",
        default=[],
        help="Only include cases that contain this tag (can be repeated).",
    )
    parser.add_argument(
        "--exclude-tag",
        action="append",
        default=[],
        help="Exclude cases that contain this tag (can be repeated).",
    )
    parser.add_argument(
        "--max-cases",
        type=int,
        default=0,
        help="Optional cap (after tag filtering): compare only the first N cases (0 = all).",
    )
    parser.add_argument("--abs-tol", type=float, default=1e-9)
    parser.add_argument("--rel-tol", type=float, default=1e-9)
    parser.add_argument(
        "--tag-abs-tol",
        action="append",
        default=[],
        help=(
            "Override numeric abs tolerance for cases that contain a tag. Format TAG=FLOAT "
            "(example: odd_coupon=1e-6). Can be repeated; the maximum across matching tags wins."
        ),
    )
    parser.add_argument(
        "--tag-rel-tol",
        action="append",
        default=[],
        help=(
            "Override numeric rel tolerance for cases that contain a tag. Format TAG=FLOAT "
            "(example: odd_coupon=1e-6). Can be repeated; the maximum across matching tags wins."
        ),
    )
    parser.add_argument(
        "--max-mismatch-rate",
        type=float,
        default=0.0,
        help="Fail if mismatches / total > this threshold (default 0).",
    )
    args = parser.parse_args()

    cases_path = Path(args.cases)
    expected_path = Path(args.expected)
    actual_path = Path(args.actual)
    report_path = Path(args.report)

    cases = _load_json(cases_path)
    expected = _load_json(expected_path)
    actual = _load_json(actual_path)

    if cases.get("schemaVersion") != 1:
        raise SystemExit(f"Unsupported cases schemaVersion: {cases.get('schemaVersion')}")
    if expected.get("schemaVersion") != 1:
        raise SystemExit(f"Unsupported expected schemaVersion: {expected.get('schemaVersion')}")
    if actual.get("schemaVersion") != 1:
        raise SystemExit(f"Unsupported actual schemaVersion: {actual.get('schemaVersion')}")

    expected_source = expected.get("source", {})
    actual_source = actual.get("source")
    if isinstance(expected_source, dict) and expected_source.get("kind") != "excel":
        raise SystemExit(
            "Expected dataset must be produced by real Excel (source.kind == 'excel'). "
            f"Got: {expected_source.get('kind')!r}"
        )

    # The repo may use a "synthetic CI baseline" pinned dataset (source.kind="excel" but with
    # source.syntheticSource metadata). Surface this explicitly in the report summary so CI tooling
    # and developers can tell at a glance whether mismatches are against real Excel or a baseline.
    expected_dataset_kind = "excel"
    expected_dataset_patch_entry_count = 0
    expected_dataset_has_patches = False
    if isinstance(expected_source, dict):
        if isinstance(expected_source.get("syntheticSource"), dict):
            expected_dataset_kind = "synthetic"
        patches = expected_source.get("patches")
        if isinstance(patches, list):
            expected_dataset_patch_entry_count = len(patches)
            expected_dataset_has_patches = expected_dataset_patch_entry_count > 0

    cases_sha = _sha256_file(cases_path)
    expected_case_set = expected.get("caseSet")
    actual_case_set = actual.get("caseSet")
    expected_sha = expected_case_set.get("sha256") if isinstance(expected_case_set, dict) else None
    actual_sha = actual_case_set.get("sha256") if isinstance(actual_case_set, dict) else None

    if isinstance(expected_sha, str) and expected_sha.lower() != cases_sha.lower():
        raise SystemExit(
            "Expected dataset caseSet.sha256 does not match cases.json. "
            f"expected={expected_sha} cases={cases_sha}"
        )

    if isinstance(actual_sha, str) and actual_sha.lower() != cases_sha.lower():
        raise SystemExit(
            "Actual dataset caseSet.sha256 does not match cases.json. "
            f"actual={actual_sha} cases={cases_sha}"
        )

    expected_results = expected.get("results", [])
    if not isinstance(expected_results, list):
        raise SystemExit("Expected dataset 'results' must be an array.")
    expected_count = expected_case_set.get("count") if isinstance(expected_case_set, dict) else None
    if isinstance(expected_count, int) and expected_count != len(expected_results):
        raise SystemExit(
            "Expected dataset caseSet.count does not match results length. "
            f"count={expected_count} results={len(expected_results)}"
        )

    actual_results = actual.get("results", [])
    if not isinstance(actual_results, list):
        raise SystemExit("Actual dataset 'results' must be an array.")
    actual_count = actual_case_set.get("count") if isinstance(actual_case_set, dict) else None
    if isinstance(actual_count, int) and actual_count != len(actual_results):
        raise SystemExit(
            "Actual dataset caseSet.count does not match results length. "
            f"count={actual_count} results={len(actual_results)}"
        )

    # Developer ergonomics: `formula-excel-oracle` is frequently run with tag filters (or `--max-cases`)
    # to keep iteration fast. If the user then runs `compare.py` without the same filters, the report
    # is dominated by "missing-actual" noise and can look like a catastrophic regression.
    #
    # When compare has no filters enabled, sanity-check that the actual dataset appears to cover the
    # full corpus before continuing.
    if (
        not args.include_tag
        and not args.exclude_tag
        and args.max_cases == 0
        and len(actual_results) != len(cases.get("cases", []))
    ):
        raise SystemExit(
            "Actual dataset does not cover the full case corpus. "
            f"cases={len(cases.get('cases', []))} actual_results={len(actual_results)}. "
            "If you generated the engine results with --include-tag/--exclude-tag or --max-cases, "
            "re-run compare.py with the same filters, or regenerate the engine results without filtering."
        )

    expected_index = _index_results(expected_results, label="Expected")
    actual_index = _index_results(actual_results, label="Actual")

    default_cfg = CompareConfig(abs_tol=args.abs_tol, rel_tol=args.rel_tol)
    tag_abs_tol = _parse_tag_tolerances(args.tag_abs_tol, flag_name="--tag-abs-tol")
    tag_rel_tol = _parse_tag_tolerances(args.tag_rel_tol, flag_name="--tag-rel-tol")

    mismatches: list[dict[str, Any]] = []
    reason_counts: dict[str, int] = {}
    tag_totals: dict[str, int] = {}
    tag_fails: dict[str, int] = {}
    missing_functions: dict[str, int] = {}
    actual_error_kinds: dict[str, int] = {}

    def _note_actual_error(actual_value: Any, *, formula_raw: str | None) -> None:
        if not (isinstance(actual_value, dict) and actual_value.get("t") == "e"):
            return
        code = actual_value.get("v")
        if not isinstance(code, str):
            return
        actual_error_kinds[code] = actual_error_kinds.get(code, 0) + 1
        if code == "#NAME?":
            for fn in _extract_function_names(formula_raw):
                missing_functions[fn] = missing_functions.get(fn, 0) + 1

    include_tags = set(args.include_tag)
    exclude_tags = set(args.exclude_tag)

    included_cases: list[dict[str, Any]] = []
    for case in cases.get("cases", []):
        case_id = case.get("id")
        if not isinstance(case_id, str):
            continue

        tags = case.get("tags", [])
        if not isinstance(tags, list):
            tags = []
        tag_set = {t for t in tags if isinstance(t, str)}

        if include_tags and not (include_tags & tag_set):
            continue
        if exclude_tags and (exclude_tags & tag_set):
            continue

        included_cases.append(case)

    matched_cases = len(included_cases)
    if args.max_cases and args.max_cases > 0:
        included_cases = included_cases[: args.max_cases]

    if args.dry_run:
        def _fmt_path(path: Path) -> str:
            raw = str(path)
            return _redact_text(raw, privacy_mode=args.privacy_mode) or ""

        print("Dry run: compare.py")
        print(f"cases: {_fmt_path(cases_path)}")
        print(f"expected: {_fmt_path(expected_path)}")
        print(f"actual: {_fmt_path(actual_path)}")
        print(f"report: {_fmt_path(report_path)}")
        print(f"cases after tag filtering: {matched_cases}")
        print(f"cases selected: {len(included_cases)}")
        return 0

    for case in included_cases:
        case_id = case["id"]
        tags = case.get("tags", [])
        if not isinstance(tags, list):
            tags = []
        tag_set = {t for t in tags if isinstance(t, str)}

        exp = expected_index.get(case_id)
        act = actual_index.get(case_id)

        mismatch_reason: str | None = None
        formula_raw = case.get("formula")
        if exp is None:
            mismatch_reason = "missing-expected"
            entry: dict[str, Any] = {
                "caseId": case_id,
                "reason": mismatch_reason,
                "formula": _redact_formula(case.get("formula"), privacy_mode=args.privacy_mode),
                "inputs": [
                    _pretty_input(i, privacy_mode=args.privacy_mode)
                    for i in case.get("inputs", [])
                ],
                "tags": sorted(tag_set),
            }
            output_cell = _maybe_nonempty_str(case.get("outputCell"))
            if output_cell is not None:
                entry["outputCell"] = output_cell
            description = _maybe_nonempty_str(case.get("description"))
            if description is not None:
                entry["description"] = description

            if isinstance(act, dict):
                # When the expected dataset is missing a case (common when new deterministic cases
                # are added to cases.json but the pinned dataset wasn't updated yet), include the
                # engine-computed value (and basic rendering metadata) to make patching/regeneration
                # easier.
                actual_value = act.get("result")
                if actual_value is not None:
                    entry["actual"] = actual_value
                    _note_actual_error(actual_value, formula_raw=formula_raw)

                actual_address = _maybe_nonempty_str(act.get("address"))
                if actual_address is not None:
                    entry["actualAddress"] = actual_address
                actual_display = _maybe_nonempty_str(act.get("displayText"))
                if actual_display is not None:
                    entry["actualDisplayText"] = actual_display

            mismatches.append(entry)
            reason_counts[mismatch_reason] = reason_counts.get(mismatch_reason, 0) + 1

        elif act is None:
            mismatch_reason = "missing-actual"
            entry = {
                "caseId": case_id,
                "reason": mismatch_reason,
                "formula": _redact_formula(case.get("formula"), privacy_mode=args.privacy_mode),
                "inputs": [
                    _pretty_input(i, privacy_mode=args.privacy_mode)
                    for i in case.get("inputs", [])
                ],
                "tags": sorted(tag_set),
                "expected": exp.get("result"),
            }
            output_cell = _maybe_nonempty_str(case.get("outputCell"))
            if output_cell is not None:
                entry["outputCell"] = output_cell
            description = _maybe_nonempty_str(case.get("description"))
            if description is not None:
                entry["description"] = description

            if isinstance(exp, dict):
                expected_address = _maybe_nonempty_str(exp.get("address"))
                if expected_address is not None:
                    entry["expectedAddress"] = expected_address
                expected_display = _maybe_nonempty_str(exp.get("displayText"))
                if expected_display is not None:
                    entry["expectedDisplayText"] = expected_display

            mismatches.append(entry)
            reason_counts[mismatch_reason] = reason_counts.get(mismatch_reason, 0) + 1

        else:
            cfg = _effective_cfg_for_tags(
                default_cfg,
                tags=tag_set,
                tag_abs_tol=tag_abs_tol,
                tag_rel_tol=tag_rel_tol,
            )
            ok, reason, mismatch_detail = _compare_value(exp.get("result"), act.get("result"), cfg)
            if not ok:
                mismatch_reason = reason
                entry: dict[str, Any] = {
                    "caseId": case_id,
                    "reason": mismatch_reason,
                    "formula": _redact_formula(case.get("formula"), privacy_mode=args.privacy_mode),
                    "inputs": [
                        _pretty_input(i, privacy_mode=args.privacy_mode)
                        for i in case.get("inputs", [])
                    ],
                    "tags": sorted(tag_set),
                    "expected": exp.get("result"),
                    "actual": act.get("result"),
                    # Record the effective numeric tolerance for this case after tag-specific
                    # overrides. This makes mismatches easier to triage when some tags (e.g.
                    # odd_coupon) intentionally use looser tolerances.
                    "absTol": cfg.abs_tol,
                    "relTol": cfg.rel_tol,
                }
                output_cell = _maybe_nonempty_str(case.get("outputCell"))
                if output_cell is not None:
                    entry["outputCell"] = output_cell
                description = _maybe_nonempty_str(case.get("description"))
                if description is not None:
                    entry["description"] = description
                if mismatch_detail is not None:
                    entry["mismatchDetail"] = mismatch_detail

                if isinstance(exp, dict):
                    expected_address = _maybe_nonempty_str(exp.get("address"))
                    if expected_address is not None:
                        entry["expectedAddress"] = expected_address
                    expected_display = _maybe_nonempty_str(exp.get("displayText"))
                    if expected_display is not None:
                        entry["expectedDisplayText"] = expected_display
                if isinstance(act, dict):
                    actual_address = _maybe_nonempty_str(act.get("address"))
                    if actual_address is not None:
                        entry["actualAddress"] = actual_address
                    actual_display = _maybe_nonempty_str(act.get("displayText"))
                    if actual_display is not None:
                        entry["actualDisplayText"] = actual_display

                exp_result = exp.get("result")
                act_result = act.get("result")
                _note_actual_error(act_result, formula_raw=formula_raw)
                if (
                    mismatch_reason == "number-mismatch"
                    and _is_number(exp_result)
                    and _is_number(act_result)
                ):
                    ev = float(exp_result.get("v"))
                    av = float(act_result.get("v"))
                    abs_diff = abs(ev - av)
                    denom = max(abs(ev), abs(av))
                    entry["absDiff"] = abs_diff
                    entry["relDiff"] = (abs_diff / denom) if denom else None

                mismatches.append(entry)
                reason_counts[mismatch_reason] = reason_counts.get(mismatch_reason, 0) + 1

        # Per-tag accounting (a case can contribute to multiple tags).
        if not tag_set:
            tag_set = {"<untagged>"}

        for t in tag_set:
            tag_totals[t] = tag_totals.get(t, 0) + 1
            if mismatch_reason is not None:
                tag_fails[t] = tag_fails.get(t, 0) + 1

    total = len(included_cases)
    mismatch_count = len(mismatches)
    mismatch_rate = (mismatch_count / total) if total else 0.0

    tag_summary: list[dict[str, Any]] = []
    for tag, tot in tag_totals.items():
        fails = tag_fails.get(tag, 0)
        passes = tot - fails
        tag_summary.append(
            {
                "tag": tag,
                "total": tot,
                "passes": passes,
                "mismatches": fails,
                "mismatchRate": (fails / tot) if tot else 0.0,
            }
        )
    tag_summary.sort(key=lambda x: (-x["mismatches"], -x["total"], x["tag"]))

    if args.privacy_mode == _PRIVACY_PRIVATE:
        # Redact tag names in the report output (defense in depth). This is intentionally applied
        # after all accounting/filtering is complete so behavior doesn't change.

        # Redact per-case mismatch tags.
        for m in mismatches:
            tags = m.get("tags")
            if isinstance(tags, list):
                out_tags: list[str] = []
                for t in tags:
                    if isinstance(t, str) and t:
                        out_tags.append(_redact_tag_name(t, privacy_mode=args.privacy_mode))
                if out_tags:
                    m["tags"] = out_tags

        # Redact tag summary labels, merging rows if redaction collapses multiple raw tags.
        tag_summary_by_name: dict[str, dict[str, Any]] = {}
        for row in tag_summary:
            raw_tag = row.get("tag")
            if not isinstance(raw_tag, str) or not raw_tag:
                continue
            redacted_tag = _redact_tag_name(raw_tag, privacy_mode=args.privacy_mode)
            existing = tag_summary_by_name.get(redacted_tag)
            if existing is None:
                tag_summary_by_name[redacted_tag] = dict(row, tag=redacted_tag)
                continue
            existing["total"] = int(existing.get("total") or 0) + int(row.get("total") or 0)
            existing["passes"] = int(existing.get("passes") or 0) + int(row.get("passes") or 0)
            existing["mismatches"] = int(existing.get("mismatches") or 0) + int(row.get("mismatches") or 0)
            tot = int(existing.get("total") or 0)
            mism = int(existing.get("mismatches") or 0)
            existing["mismatchRate"] = (mism / tot) if tot else 0.0
        tag_summary = list(tag_summary_by_name.values())
        tag_summary.sort(key=lambda x: (-x["mismatches"], -x["total"], x["tag"]))

        # Redact tag tolerance mapping keys (user-provided CLI tags can include custom namespaces).
        def _redact_tag_tolerance_map(data: dict[str, float]) -> dict[str, float]:
            out: dict[str, float] = {}
            for k, v in data.items():
                if not isinstance(k, str) or not k:
                    continue
                key = _redact_tag_name(k, privacy_mode=args.privacy_mode)
                prev = out.get(key)
                if prev is None or float(v) > float(prev):
                    out[key] = float(v)
            return out

        tag_abs_tol = _redact_tag_tolerance_map(tag_abs_tol)
        tag_rel_tol = _redact_tag_tolerance_map(tag_rel_tol)

    if args.privacy_mode == _PRIVACY_PRIVATE and missing_functions:
        redacted_counts: dict[str, int] = {}
        for fn, cnt in missing_functions.items():
            key = _redact_function_name(fn, privacy_mode=args.privacy_mode)
            redacted_counts[key] = redacted_counts.get(key, 0) + int(cnt)
        missing_functions = redacted_counts

    top_missing_functions = [
        {"name": k, "count": v}
        for k, v in sorted(missing_functions.items(), key=lambda kv: (-kv[1], kv[0]))
    ][:20]
    top_actual_error_kinds = [
        {"code": k, "count": v}
        for k, v in sorted(actual_error_kinds.items(), key=lambda kv: (-kv[1], kv[0]))
    ][:20]

    summary: dict[str, Any] = {
        "totalCases": total,
        "includeTags": sorted(
            _redact_tag_name(t, privacy_mode=args.privacy_mode) for t in include_tags
        ),
        "excludeTags": sorted(
            _redact_tag_name(t, privacy_mode=args.privacy_mode) for t in exclude_tags
        ),
        "maxCases": args.max_cases,
        "absTol": args.abs_tol,
        "relTol": args.rel_tol,
        "tagAbsTol": tag_abs_tol,
        "tagRelTol": tag_rel_tol,
        "mismatches": mismatch_count,
        "mismatchRate": mismatch_rate,
        "maxMismatchRate": args.max_mismatch_rate,
        "reasonCounts": dict(sorted(reason_counts.items(), key=lambda kv: (-kv[1], kv[0]))),
        "tagSummary": tag_summary,
        "topMissingFunctions": top_missing_functions,
        "topActualErrorKinds": top_actual_error_kinds,
        "casesSha256": cases_sha,
        # Make reports self-contained: consumers (CI artifacts, local debugging) should be able
        # to see exactly which datasets were compared without having to reconstruct CLI args.
        "casesPath": str(cases_path),
        "expectedPath": str(expected_path),
        "actualPath": str(actual_path),
        # Expected dataset provenance.
        "expectedDatasetKind": expected_dataset_kind,
        "expectedDatasetHasPatches": expected_dataset_has_patches,
        "expectedDatasetPatchEntryCount": expected_dataset_patch_entry_count,
    }
    summary = _redact_paths_in_obj(summary, privacy_mode=args.privacy_mode)

    report = {
        "schemaVersion": 1,
        "summary": summary,
        "expectedSource": _redact_paths_in_obj(expected.get("source"), privacy_mode=args.privacy_mode),
        "actualSource": _redact_paths_in_obj(actual_source, privacy_mode=args.privacy_mode),
        "mismatches": mismatches,
    }
    report = _redact_error_details_in_obj(report, privacy_mode=args.privacy_mode)
    report = _redact_string_values_in_obj(report, privacy_mode=args.privacy_mode)

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", encoding="utf-8", newline="\n") as f:
        json.dump(report, f, ensure_ascii=False, indent=2, sort_keys=False)
        f.write("\n")

    # Human-friendly summary (stdout) for CI/dev ergonomics.
    print(f"Excel compatibility: {total} cases, {mismatch_count} mismatches ({mismatch_rate:.4%})")
    if tag_summary:
        print("")
        print("Tag summary (mismatches/total):")
        # Print tags with failures; if none, print the largest tags for context.
        interesting = [t for t in tag_summary if t["mismatches"] > 0]
        if not interesting:
            interesting = tag_summary[: min(10, len(tag_summary))]
        for row in interesting[: min(25, len(interesting))]:
            print(
                f"  {row['tag']}: {row['mismatches']}/{row['total']} ({row['mismatchRate']:.2%})"
            )

    if top_missing_functions:
        print("")
        print("Top missing functions (mismatches where actual is #NAME?):")
        for row in top_missing_functions[: min(10, len(top_missing_functions))]:
            print(f"  {row['name']}: {row['count']}")

    if top_actual_error_kinds:
        print("")
        print("Top actual error kinds (in mismatches):")
        for row in top_actual_error_kinds[: min(10, len(top_actual_error_kinds))]:
            print(f"  {row['code']}: {row['count']}")

    # Exit code based on threshold.
    if mismatch_rate > args.max_mismatch_rate:
        sys.stderr.write(
            f"Excel compatibility mismatch rate {mismatch_rate:.4%} exceeded threshold {args.max_mismatch_rate:.4%}\n"
        )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
