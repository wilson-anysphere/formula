from __future__ import annotations

import base64
import hashlib
import json
import os
import subprocess
import tempfile
from functools import lru_cache
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Iterator


def sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: Any) -> None:
    ensure_dir(path.parent)
    text = json.dumps(data, indent=2, sort_keys=True) + "\n"
    # Write atomically to avoid leaving partially-written JSON if the process is interrupted.
    with tempfile.NamedTemporaryFile(
        "w",
        encoding="utf-8",
        delete=False,
        dir=path.parent,
        prefix=path.name + ".",
        suffix=".tmp",
    ) as f:
        f.write(text)
        tmp_name = f.name
    os.replace(tmp_name, path)


@dataclass(frozen=True)
class WorkbookInput:
    """A workbook blob plus a stable display name (no local paths)."""

    display_name: str
    data: bytes


def read_workbook_input(path: Path, *, fernet_key: str | None = None) -> WorkbookInput:
    """Read a workbook blob from disk.

    Supports:
    - raw `.xlsx`/`.xlsm`/`.xlsb` files
    - base64-encoded `*.xlsx.b64`/`*.xlsm.b64`/`*.xlsb.b64` (text) for embedding small public fixtures
    - encrypted `*.enc` files (Fernet) for private corpus storage
    """

    data = path.read_bytes()
    display_name = path.name

    if display_name.casefold().endswith(".b64"):
        # Base64 fixtures are stored as text so they can live in git.
        # `validate=True` rejects whitespace/newlines, so strip them first to support
        # checked-in fixtures with trailing newlines.
        data = base64.b64decode(b"".join(data.split()), validate=True)
        display_name = display_name[: -len(".b64")]

    if display_name.casefold().endswith(".enc"):
        if not fernet_key:
            raise ValueError(
                f"{path} looks encrypted ('.enc') but no `fernet_key` was provided."
            )
        from cryptography.fernet import Fernet  # local import to keep dependency optional

        f = Fernet(fernet_key.encode("utf-8"))
        data = f.decrypt(data)
        display_name = display_name[: -len(".enc")]

    return WorkbookInput(display_name=display_name, data=data)


def iter_workbook_paths(corpus_dir: Path, *, include_xlsb: bool = False) -> Iterator[Path]:
    """Yield candidate workbook paths under `corpus_dir`.

    By default only XLSX/XLSM workbooks are yielded. Pass `include_xlsb=True` to also include
    `.xlsb` (and `.xlsb.b64`/`.xlsb.enc`) fixtures.
    """

    for path in sorted(corpus_dir.rglob("*")):
        if not path.is_file():
            continue
        name = path.name.lower()
        endings = (".xlsx", ".xlsm", ".xlsx.b64", ".xlsm.b64", ".xlsx.enc", ".xlsm.enc")
        if include_xlsb:
            endings = endings + (".xlsb", ".xlsb.b64", ".xlsb.enc")

        if name.endswith(endings):
            yield path


@lru_cache(maxsize=1)
def _local_git_commit_sha() -> str | None:
    """Best-effort git HEAD SHA for local runs.

    Cached so per-workbook triage doesn't spawn `git` repeatedly.
    """

    try:
        root = Path(__file__).resolve().parents[2]
        proc = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=root,
            capture_output=True,
            text=True,
            check=True,
        )
        out = (proc.stdout or "").strip()
        return out or None
    except Exception:  # noqa: BLE001
        return None


def github_commit_sha() -> str | None:
    sha = os.environ.get("GITHUB_SHA")
    if sha:
        return sha
    # Local runs: best-effort fallback to the current git commit so trend files can be used
    # deterministically outside of GitHub Actions. This is intentionally non-fatal: if `git`
    # isn't available (or we're not in a worktree), return None.
    return _local_git_commit_sha()


def github_run_url() -> str | None:
    server = os.environ.get("GITHUB_SERVER_URL")
    repo = os.environ.get("GITHUB_REPOSITORY")
    run_id = os.environ.get("GITHUB_RUN_ID")
    if server and repo and run_id:
        return f"{server}/{repo}/actions/runs/{run_id}"
    return None
