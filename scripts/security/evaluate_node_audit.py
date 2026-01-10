#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


HIGH_SEVERITIES = {"high", "critical"}


def _load_allowlist(path: Path | None) -> set[str]:
    if path is None or not path.exists():
        return set()
    allowlisted: set[str] = set()
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        allowlisted.add(line.split()[0])
    return allowlisted


def _extract_ids(text: str) -> set[str]:
    ids: set[str] = set()
    for pat in (r"GHSA-[0-9a-zA-Z]{4}-[0-9a-zA-Z]{4}-[0-9a-zA-Z]{4}", r"CVE-\d{4}-\d{4,}"):
        ids.update(re.findall(pat, text))
    # npm advisory URLs contain ".../advisories/<id>"
    m = re.search(r"/advisories/(\d+)", text)
    if m:
        ids.add(m.group(1))
    return ids


@dataclass(frozen=True)
class Advisory:
    ids: set[str]
    severity: str
    title: str | None
    package: str | None
    url: str | None


def _iter_v2_advisories(doc: dict[str, Any]) -> Iterable[Advisory]:
    """
    npm audit JSON v2 structure:
      { "auditReportVersion": 2, "vulnerabilities": { ... }, "metadata": { ... } }
    """
    vulns = doc.get("vulnerabilities")
    if not isinstance(vulns, dict):
        return

    for pkg_name, entry in vulns.items():
        if not isinstance(entry, dict):
            continue
        via = entry.get("via")
        if not isinstance(via, list):
            continue
        for item in via:
            if isinstance(item, str):
                # Transitive reference, not an advisory object
                continue
            if not isinstance(item, dict):
                continue

            severity = item.get("severity")
            if not isinstance(severity, str):
                continue
            url = item.get("url") if isinstance(item.get("url"), str) else None
            title = item.get("title") if isinstance(item.get("title"), str) else None

            ids: set[str] = set()
            if url:
                ids |= _extract_ids(url)
            # Some npm advisory objects include "cwe"/"cves"/etc in different forms.
            for k in ("cves", "cwe", "name", "source"):
                if k in item and isinstance(item[k], str):
                    ids |= _extract_ids(item[k])
                if k in item and isinstance(item[k], int):
                    ids.add(str(item[k]))

            yield Advisory(
                ids=ids if ids else {pkg_name},
                severity=severity.lower().strip(),
                title=title,
                package=pkg_name,
                url=url,
            )


def _iter_v1_advisories(doc: dict[str, Any]) -> Iterable[Advisory]:
    """
    npm audit JSON v1 structure:
      { "advisories": { "<id>": { ... } }, "metadata": { ... } }
    """
    advisories = doc.get("advisories")
    if not isinstance(advisories, dict):
        return

    for _, adv in advisories.items():
        if not isinstance(adv, dict):
            continue
        severity = adv.get("severity")
        if not isinstance(severity, str):
            continue
        title = adv.get("title") if isinstance(adv.get("title"), str) else None
        module_name = adv.get("module_name") if isinstance(adv.get("module_name"), str) else None

        ids: set[str] = set()
        if isinstance(adv.get("github_advisory_id"), str):
            ids |= _extract_ids(adv["github_advisory_id"])
        if isinstance(adv.get("cves"), list):
            for cve in adv["cves"]:
                if isinstance(cve, str):
                    ids |= _extract_ids(cve)
        if isinstance(adv.get("url"), str):
            ids |= _extract_ids(adv["url"])

        # The advisory "id" is an npm advisory numeric ID.
        if isinstance(adv.get("id"), int):
            ids.add(str(adv["id"]))

        yield Advisory(
            ids=ids if ids else {module_name or "unknown"},
            severity=severity.lower().strip(),
            title=title,
            package=module_name,
            url=adv.get("url") if isinstance(adv.get("url"), str) else None,
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="Evaluate pnpm/npm audit output for CI policy.")
    parser.add_argument("--input", required=True, type=Path, help="Path to audit JSON output")
    parser.add_argument("--allowlist", type=Path, default=None, help="Allowlist file (one ID per line)")
    parser.add_argument("--output", type=Path, required=True, help="Where to write a JSON policy summary")
    args = parser.parse_args()

    allowlisted = _load_allowlist(args.allowlist)

    try:
        raw = args.input.read_text(encoding="utf-8")
        doc = json.loads(raw) if raw.strip() else {}
    except Exception as exc:  # noqa: BLE001
        args.output.write_text(
            json.dumps(
                {
                    "ok": False,
                    "error": f"Failed to parse audit JSON: {exc}",
                    "input": str(args.input),
                },
                indent=2,
                sort_keys=True,
            )
            + "\n",
            encoding="utf-8",
        )
        print(f"node-audit: ERROR parsing JSON ({exc})", file=sys.stderr)
        return 1

    advisories: list[Advisory] = []
    if isinstance(doc, dict) and doc.get("auditReportVersion") == 2:
        advisories.extend(list(_iter_v2_advisories(doc)))
    if isinstance(doc, dict) and "advisories" in doc:
        advisories.extend(list(_iter_v1_advisories(doc)))

    high_or_critical: list[Advisory] = []
    ignored: list[Advisory] = []

    for adv in advisories:
        if adv.severity not in HIGH_SEVERITIES:
            continue
        if adv.ids & allowlisted:
            ignored.append(adv)
            continue
        high_or_critical.append(adv)

    ok = len(high_or_critical) == 0

    summary = {
        "ok": ok,
        "input": str(args.input),
        "allowlist": str(args.allowlist) if args.allowlist else None,
        "total_advisories": len(advisories),
        "total_ignored": len(ignored),
        "total_high_or_critical": len(high_or_critical),
        "high_or_critical": [
            {
                "ids": sorted(a.ids),
                "severity": a.severity,
                "package": a.package,
                "title": a.title,
                "url": a.url,
            }
            for a in high_or_critical
        ],
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    label = "OK" if ok else "FAIL"
    print(
        f"node audit policy: {label} "
        f"(total={summary['total_advisories']}, "
        f"ignored={summary['total_ignored']}, "
        f"high_or_critical={summary['total_high_or_critical']})"
    )
    if not ok:
        for a in high_or_critical:
            ids = ",".join(sorted(a.ids))
            pkg = a.package or "unknown"
            title = a.title or "unknown"
            print(f"  - {ids} package={pkg} severity={a.severity} title={title}")

    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())

