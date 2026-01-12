# Sample Hello (reference extension)

This is a small, runnable reference extension used by the repository’s integration tests and
marketplace packaging tests.

## Features

- **Commands**
  - `sampleHello.sumSelection`: sums the current selection, writes the sum to cell (row=2, col=0),
    and shows `Sum: <sum>`.
  - `sampleHello.openPanel`: creates `sampleHello.panel` and sets its HTML.
  - `sampleHello.fetchText <url>`: fetches text (permission-gated) and shows `Fetched: <text>`.
  - `sampleHello.copySumToClipboard`: sums selection and writes it to the clipboard.
  - `sampleHello.showGreeting`: reads `sampleHello.greeting` from config and shows `Greeting: ...`.
- **Panel**
  - HTML contains the string `Sample Hello Panel`
  - Responds to `{ type: "ping" }` messages with `{ type: "pong" }`
  - Auto-creates/renders on `onViewActivated` for `sampleHello.panel`
- **Custom function**
  - `SAMPLEHELLO_DOUBLE(value)` returns `value * 2`

## Build workflow

`src/extension.js` is the source of truth.

`dist/extension.js` is the Node (CommonJS) entrypoint referenced by `package.json` (`main`).

`dist/extension.mjs` is an ES module build that can be used by the browser extension host
(`package.json` `browser`/`module` fields).

## Desktop/Tauri runtime notes (WebView extension host)

In Formula Desktop (Tauri/WebView), extensions run inside a **Web Worker** (`BrowserExtensionHost`) with a best-effort
sandbox (permission-gated `fetch`/`WebSocket`, no `XMLHttpRequest`, etc).

Extension panels (`contributes.panels` / `formula.ui.createPanel`) are rendered in a sandboxed `<iframe>` with a very
restrictive CSP (notably `connect-src 'none'`), so panel HTML:

- cannot make network requests directly
- cannot load remote scripts
- cannot run inline `<script>` blocks (scripts must be loaded via `data:`/`blob:` URLs)

Panels should communicate with the extension worker via `postMessage`.

Marketplace-installed extensions are loaded from in-memory `blob:`/`data:` module URLs, so the browser entrypoint should
be a **single-file ESM bundle** (no relative imports). This sample’s `dist/extension.mjs` is built that way.

To regenerate:

```bash
pnpm extension:build extensions/sample-hello
```

CI runs a node:test (`build.test.js`) that fails if `dist/extension.js` is out of date.
