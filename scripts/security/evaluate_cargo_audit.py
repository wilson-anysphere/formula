#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
import math


HIGH_SEVERITIES = {"high", "critical"}


def _load_allowlist(path: Path | None) -> set[str]:
    if path is None or not path.exists():
        return set()
    allowlisted: set[str] = set()
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        # Allow inline comments: "RUSTSEC-.... # reason"
        allowlisted.add(line.split()[0])
    return allowlisted


def _as_float(value: Any) -> float | None:
    try:
        if value is None:
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _severity_is_high_or_critical(advisory: dict[str, Any]) -> bool:
    """
    cargo-audit JSON contains advisory metadata but not always a severity.

    Policy:
    - If severity is present and is "high"/"critical" => fail.
    - If severity is present and is lower => ok.
    - If severity is missing but a numeric CVSS score is present => fail when score >= 7.0.
    - If we can't determine severity => treat as high (fail-safe).
    """

    severity = advisory.get("severity")
    if isinstance(severity, str):
        s = severity.lower().strip()
        if s in HIGH_SEVERITIES:
            return True
        return False

    cvss = advisory.get("cvss")
    if isinstance(cvss, dict):
        score = _as_float(cvss.get("score"))
        if score is not None:
            return score >= 7.0
        vector = cvss.get("vector") if isinstance(cvss.get("vector"), str) else None
        if vector:
            score = _cvss_v3_base_score(vector)
            if score is not None:
                return score >= 7.0
    if isinstance(cvss, str):
        score = _cvss_v3_base_score(cvss)
        if score is not None:
            return score >= 7.0

    # Unknown severity -> fail-safe
    return True


def _cvss_v3_base_score(vector: str) -> float | None:
    """
    Minimal CVSS v3.0/v3.1 base score calculator.

    RustSec advisories often store CVSS as a vector string rather than a precomputed
    score. We compute the base score to map onto severity thresholds:
      - High:     7.0 - 8.9
      - Critical: 9.0 - 10.0
    """

    if not vector:
        return None
    v = vector.strip()
    if v.startswith("CVSS:3."):
        # Drop the prefix "CVSS:3.1/"
        parts = v.split("/", 1)
        if len(parts) == 2:
            v = parts[1]

    metrics: dict[str, str] = {}
    for part in v.split("/"):
        if ":" not in part:
            continue
        k, val = part.split(":", 1)
        metrics[k] = val

    required = {"AV", "AC", "PR", "UI", "S", "C", "I", "A"}
    if not required.issubset(metrics.keys()):
        return None

    av_map = {"N": 0.85, "A": 0.62, "L": 0.55, "P": 0.2}
    ac_map = {"L": 0.77, "H": 0.44}
    ui_map = {"N": 0.85, "R": 0.62}
    cia_map = {"H": 0.56, "L": 0.22, "N": 0.0}

    s = metrics["S"]
    if s not in {"U", "C"}:
        return None

    pr_map_u = {"N": 0.85, "L": 0.62, "H": 0.27}
    pr_map_c = {"N": 0.85, "L": 0.68, "H": 0.5}
    pr_map = pr_map_c if s == "C" else pr_map_u

    try:
        av = av_map[metrics["AV"]]
        ac = ac_map[metrics["AC"]]
        pr = pr_map[metrics["PR"]]
        ui = ui_map[metrics["UI"]]
        c = cia_map[metrics["C"]]
        i = cia_map[metrics["I"]]
        a = cia_map[metrics["A"]]
    except KeyError:
        return None

    impact_sub = 1 - (1 - c) * (1 - i) * (1 - a)
    if s == "U":
        impact = 6.42 * impact_sub
    else:
        impact = 7.52 * (impact_sub - 0.029) - 3.25 * pow(impact_sub - 0.02, 15)

    exploitability = 8.22 * av * ac * pr * ui

    if impact <= 0:
        return 0.0

    score = impact + exploitability
    if s == "C":
        score *= 1.08
    score = min(score, 10.0)

    # Round up to one decimal place.
    return math.ceil(score * 10) / 10.0


@dataclass(frozen=True)
class Finding:
    advisory_id: str
    package: str | None
    severity: str | None


def _iter_findings(doc: dict[str, Any]) -> Iterable[Finding]:
    vulns = doc.get("vulnerabilities") or {}
    items = vulns.get("list") or []
    if not isinstance(items, list):
        return []
    for item in items:
        if not isinstance(item, dict):
            continue
        advisory = item.get("advisory") or {}
        if not isinstance(advisory, dict):
            continue
        advisory_id = advisory.get("id")
        if not isinstance(advisory_id, str):
            continue
        pkg = advisory.get("package")
        package_name = None
        if isinstance(pkg, dict):
            package_name = pkg.get("name") if isinstance(pkg.get("name"), str) else None
        severity = advisory.get("severity") if isinstance(advisory.get("severity"), str) else None
        yield Finding(advisory_id=advisory_id, package=package_name, severity=severity)


def main() -> int:
    parser = argparse.ArgumentParser(description="Evaluate cargo audit output for CI policy.")
    parser.add_argument("--input", required=True, type=Path, help="Path to cargo audit JSON output")
    parser.add_argument("--allowlist", type=Path, default=None, help="Allowlist file (one RUSTSEC-* per line)")
    parser.add_argument("--output", type=Path, required=True, help="Where to write a JSON policy summary")
    args = parser.parse_args()

    allowlisted = _load_allowlist(args.allowlist)

    try:
        raw = args.input.read_text(encoding="utf-8")
        doc = json.loads(raw) if raw.strip() else {}
    except Exception as exc:  # noqa: BLE001 - surface error in CI output
        args.output.write_text(
            json.dumps(
                {
                    "ok": False,
                    "error": f"Failed to parse cargo audit JSON: {exc}",
                    "input": str(args.input),
                },
                indent=2,
                sort_keys=True,
            )
            + "\n",
            encoding="utf-8",
        )
        print(f"cargo-audit: ERROR parsing JSON ({exc})", file=sys.stderr)
        return 1

    findings = list(_iter_findings(doc))
    high_or_critical: list[Finding] = []
    ignored: list[Finding] = []

    # Re-iterate with full advisory data for robust severity evaluation.
    items = ((doc.get("vulnerabilities") or {}).get("list") or [])
    if isinstance(items, list):
        for item in items:
            if not isinstance(item, dict):
                continue
            advisory = item.get("advisory") or {}
            if not isinstance(advisory, dict):
                continue
            advisory_id = advisory.get("id")
            if not isinstance(advisory_id, str):
                continue
            pkg = advisory.get("package")
            package_name = None
            if isinstance(pkg, dict):
                package_name = pkg.get("name") if isinstance(pkg.get("name"), str) else None
            sev = advisory.get("severity") if isinstance(advisory.get("severity"), str) else None
            finding = Finding(advisory_id=advisory_id, package=package_name, severity=sev)

            if advisory_id in allowlisted:
                ignored.append(finding)
                continue
            if _severity_is_high_or_critical(advisory):
                high_or_critical.append(finding)

    ok = len(high_or_critical) == 0
    summary = {
        "ok": ok,
        "input": str(args.input),
        "allowlist": str(args.allowlist) if args.allowlist else None,
        "total_vulnerabilities": len(findings),
        "total_ignored": len(ignored),
        "total_high_or_critical": len(high_or_critical),
        "high_or_critical": [
            {
                "id": f.advisory_id,
                "package": f.package,
                "severity": f.severity,
            }
            for f in high_or_critical
        ],
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    label = "OK" if ok else "FAIL"
    print(
        f"cargo-audit policy: {label} "
        f"(total={summary['total_vulnerabilities']}, "
        f"ignored={summary['total_ignored']}, "
        f"high_or_critical={summary['total_high_or_critical']})"
    )
    if not ok:
        for f in high_or_critical:
            pkg = f" ({f.package})" if f.package else ""
            sev = f.severity or "unknown"
            print(f"  - {f.advisory_id}{pkg} severity={sev}")

    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
