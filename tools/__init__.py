"""Repository tooling namespace package.

The production application lives elsewhere; this package exists so the CI scripts in `tools/**`
can use stable imports when run as modules (e.g. `python -m tools.corpus.triage`).
"""

