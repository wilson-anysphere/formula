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

`dist/extension.js` is the built entrypoint referenced by `package.json` (`main`).

To regenerate:

```bash
node extensions/sample-hello/build.js
```

CI runs a node:test (`build.test.js`) that fails if `dist/extension.js` is out of date.

