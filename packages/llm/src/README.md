# `packages/llm` (Cursor backend)

This package provides a small, dependency-free LLM client that talks to the **Cursor-managed backend**.

## Cursor-only constraints

- No provider selection (Cursor controls routing/model choice).
- No user-supplied API keys (the `apiKey` option is not supported and will throw; legacy provider env vars are **not** used by the Cursor client).
- Callers must inject auth using **request headers**.

## Backend protocol

The Cursor backend is expected to expose a **Chat Completions-compatible** endpoint:

- `POST /chat/completions`
- `messages` + `tools` + `tool_calls` follow the Chat Completions tool-calling format.

## Auth injection

Authentication is supplied by the embedding environment (desktop app, web app, etc) and injected via:

- `getAuthHeaders(): Record<string,string> | Promise<Record<string,string>>` (merged into request headers)
- `authToken` (adds `Authorization: Bearer <token>`)

## Minimal usage

### `createLLMClient()`

```ts
import { createLLMClient } from "./index.js";

// Uses the Cursor backend. Auth is expected to be handled by the embedding
// environment (for example, session cookies in a browser runtime).
const client = createLLMClient();

const res = await client.chat({
  messages: [{ role: "user", content: "Hello" }],
});
console.log(res.message.content);
```

### `new CursorLLMClient(...)`

```ts
import { CursorLLMClient } from "./index.js";

const client = new CursorLLMClient({
  baseUrl: "https://api.cursor.sh/v1",
  authToken: "cursor-managed-token",
});
```

### Tool result serialization (bounded)

```ts
import { serializeToolResultForModel } from "./index.js";

const toolResultForModel = serializeToolResultForModel({
  toolCall: { id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:Z1000" } },
  result: someToolResult,
  maxChars: 20_000,
});
```
