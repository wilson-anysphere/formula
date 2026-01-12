# XLS (`.xls`) note/comment import hardening

Legacy Excel `.xls` “notes” (cell comments) are stored in BIFF worksheet substreams using a trio of
records:

- `NOTE` (author + cell anchor + object id)
- `OBJ` (drawing object container; used to discover the object id for the note shape)
- `TXO` (+ `CONTINUE`) (the note text payload)

In this repo, the NOTE/OBJ/TXO parsing lives in:

- `crates/formula-xls/src/biff/comments.rs`

## Task tracker (follow-ups)

This section exists to avoid duplicated follow-up work across the task queue.

### Task 135 — Robust NOTE author parsing (includes NUL stripping)

**Status: closed (implemented).**

Goal: make NOTE author parsing resilient to malformed/producer-divergent BIFF payloads.
This task **supersedes Task 140** (embedded NUL stripping) to avoid duplicated work.

Scope:

- NOTE **author string** parsing only (not the TXO text payload).
- Must include **embedded NUL (`\0`) stripping** for author names.
- If a producer stores the author using an unexpected string encoding/layout, prefer **best-effort**
  decode rather than dropping the note entirely (incl. BIFF8 `XLUnicodeString` fallback).

Implementation note:

- Implemented in `crates/formula-xls/src/biff/comments.rs` (`parse_note_record`):
  - parses `ShortXLUnicodeString` / BIFF5 short strings
  - falls back to BIFF8 `XLUnicodeString` when payload layout suggests it
  - strips embedded `\0` characters from the decoded author string

Merged/closed tasks:

- Task 140 — Strip embedded NULs in NOTE author: **closed** (duplicate/subset of Task 135)
