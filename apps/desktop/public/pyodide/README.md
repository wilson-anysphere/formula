# Pyodide assets (local hosting)

The desktop Vite dev server enables `crossOriginIsolated` (COOP/COEP) so the
Pyodide-based Python runtime can run Pyodide in a Worker and use `SharedArrayBuffer`
for synchronous spreadsheet RPC.

In that mode, we **self-host** the Pyodide distribution files under the Vite
origin (COEP-friendly same-origin assets):

`/pyodide/v0.25.1/full/*`

The files are downloaded on-demand by:

```bash
pnpm -C apps/desktop dev
# or (download without starting Vite):
pnpm -C apps/desktop pyodide:setup
```

They are intentionally not checked into git (see `.gitignore` under
`v0.25.1/full/`).
