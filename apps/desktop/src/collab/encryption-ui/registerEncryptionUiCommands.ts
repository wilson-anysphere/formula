import type { SpreadsheetApp } from "../../app/spreadsheetApp";
import type { CommandRegistry } from "../../extensions/commandRegistry.js";
import { showInputBox, showQuickPick, showToast } from "../../extensions/ui.js";
import type { Range } from "../../selection/types";
import { rangeToA1 } from "../../selection/a1";

import { base64ToBytes, bytesToBase64, isEncryptedCellPayload } from "@formula/collab-encryption";
import { createEncryptionPolicyFromDoc, type EncryptedRangeManager } from "@formula/collab-encrypted-ranges";
import { serializeEncryptionKeyExportString, parseEncryptionKeyExportString } from "./keyExportFormat";

const COMMAND_CATEGORY = "Collaboration";

function normalizeRange(range: Range): { startRow: number; endRow: number; startCol: number; endCol: number } {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

function normalizeSheetNameForCompare(name: string): string {
  try {
    return String(name ?? "").normalize("NFKC").toUpperCase();
  } catch {
    return String(name ?? "").toUpperCase();
  }
}

function rangesIntersect(
  a: { startRow: number; startCol: number; endRow: number; endCol: number },
  b: { startRow: number; startCol: number; endRow: number; endCol: number },
): boolean {
  if (a.endRow < b.startRow || b.endRow < a.startRow) return false;
  if (a.endCol < b.startCol || b.endCol < a.startCol) return false;
  return true;
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (!(a instanceof Uint8Array) || !(b instanceof Uint8Array)) return false;
  if (a.byteLength !== b.byteLength) return false;
  for (let i = 0; i < a.byteLength; i += 1) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

function keyIdFromEncryptedCellPayload(session: any, cell: { sheetId: string; row: number; col: number }): string | null {
  try {
    const cells = session?.cells;
    if (!cells || typeof cells.get !== "function") return null;
    const sheetId = String(cell.sheetId ?? "").trim();
    const row = Number(cell.row);
    const col = Number(cell.col);
    if (!sheetId || !Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) return null;

    const keys: string[] = [`${sheetId}:${row}:${col}`, `${sheetId}:${row},${col}`];
    const defaultSheetId = String((session as any)?.defaultSheetId ?? "").trim();
    if (defaultSheetId && defaultSheetId === sheetId) {
      keys.push(`r${row}c${col}`);
    }

    for (const key of keys) {
      const raw = cells.get(key);
      if (!raw || typeof raw !== "object" || typeof (raw as any).get !== "function") continue;
      const enc = (raw as any).get("enc");
      if (enc == null) continue;
      if (!isEncryptedCellPayload(enc)) continue;
      const keyId = String((enc as any).keyId ?? "").trim();
      if (keyId) return keyId;
    }
  } catch {
    // ignore
  }
  return null;
}

function roleCanEncrypt(role: string | null | undefined): boolean {
  return role === "owner" || role === "admin" || role === "editor";
}

function randomKeyId(): string {
  const cryptoObj: any = globalThis.crypto as any;
  if (cryptoObj?.randomUUID) return cryptoObj.randomUUID();
  return `enc_${Math.random().toString(16).slice(2)}_${Date.now()}`;
}

function randomKeyBytes(): Uint8Array {
  const cryptoObj: any = globalThis.crypto as any;
  if (!cryptoObj?.getRandomValues) {
    throw new Error("WebCrypto is required for encryption key generation (crypto.getRandomValues missing)");
  }
  return cryptoObj.getRandomValues(new Uint8Array(32));
}

async function tryCopyToClipboard(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}

function getEncryptionManager(app: SpreadsheetApp): EncryptedRangeManager | null {
  return app.getEncryptedRangeManager();
}

export function registerEncryptionUiCommands(opts: { commandRegistry: CommandRegistry; app: SpreadsheetApp }): void {
  const { commandRegistry, app } = opts;

  commandRegistry.registerBuiltinCommand(
    "collab.encryptSelectedRange",
    "Encrypt selected range…",
    async () => {
      const session = app.getCollabSession();
      if (!session) {
        showToast("This command requires collaboration mode.", "warning");
        return;
      }
      const role = session.getRole();
      if (!roleCanEncrypt(role)) {
        showToast("You must have an editor role to encrypt ranges.", "warning");
        return;
      }

      const manager = getEncryptionManager(app);
      if (!manager) {
        showToast("Encryption is not available for this workbook.", "error");
        return;
      }

      const ranges = app.getSelectionRanges();
      if (!ranges.length) {
        showToast("Select a range to encrypt.", "warning");
        return;
      }
      if (ranges.length > 1) {
        showToast("Encrypt range currently supports a single rectangular selection.", "warning");
        return;
      }

      const sheetId = app.getCurrentSheetId();
      const sheetName = app.getCurrentSheetDisplayName();
      const range = normalizeRange(ranges[0]!);
      const a1 = rangeToA1(range);

      const docId = session.doc.guid;
      const keyStore = app.getCollabEncryptionKeyStore();
      if (!keyStore) {
        showToast("Encryption is not available for this workbook.", "error");
        return;
      }

      // Ensure the encrypted range metadata is readable before we prompt for a key id or generate/store
      // key material. If the doc contains an unknown `metadata.encryptedRanges` schema, the manager
      // will throw; fail early so we don't create orphaned keys.
      try {
        manager.list();
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        showToast(`Encrypted range metadata is in an unsupported format: ${message}`, "error");
        return;
      }

      const rawKeyId = await showInputBox({
        prompt: `Encrypt ${sheetName}!${a1} – Key ID`,
        value: randomKeyId(),
        placeHolder: "Key ID (for example: team-budget-q1)",
        okLabel: "Next",
      });
      if (!rawKeyId) return;
      const keyId = String(rawKeyId).trim();
      if (!keyId) {
        showToast("Key ID is required.", "warning");
        return;
      }

      // If the key id already exists, reuse the stored key bytes rather than overwriting
      // them (overwriting would make previously-encrypted cells unrecoverable).
      let existingKeyBytes: Uint8Array | null = null;
      let existingStoredKeyId: string | null = null;
      let existingKeyLookupFailed = false;
      try {
        const cached = keyStore.getCachedKey(docId, keyId);
        if (cached?.keyBytes instanceof Uint8Array) {
          existingKeyBytes = cached.keyBytes;
          existingStoredKeyId = cached.keyId;
        } else {
          const entry = await keyStore.get(docId, keyId);
          if (entry) {
            existingKeyBytes = base64ToBytes(entry.keyBytesBase64);
            existingStoredKeyId = entry.keyId;
          }
        }
      } catch {
        // Best-effort: if key lookup fails, proceed as if missing. This can happen when the
        // persistent key store is unavailable; warn the user in the confirmation prompt since
        // proceeding could overwrite an existing key id.
        existingKeyLookupFailed = true;
      }
      const isReusingKey = existingKeyBytes != null;
      const displayKeyId = existingStoredKeyId ?? keyId;

      // Safety: if the key id is already used in the document policy but we don't have key bytes,
      // we must not generate new key material for the same id (it would conflict with the existing key
      // and could make collaborators' encrypted cells unreadable).
      if (!existingKeyBytes) {
        try {
          const existingRanges = manager.list();
          const keyIdInUse = existingRanges.some((r) => String(r.keyId ?? "").trim() === keyId);
          if (keyIdInUse) {
            showToast(
              `Key ID "${keyId}" is already used by an encrypted range. Import the key first (or choose a new Key ID) to avoid key conflicts.`,
              "warning",
            );
            return;
          }
        } catch (err) {
          // If we can't read encrypted range metadata, we must not generate/store key material
          // (it could create orphaned keys or key conflicts).
          const message = err instanceof Error ? err.message : String(err);
          showToast(`Encrypted range metadata is in an unsupported format: ${message}`, "error");
          return;
        }

        // Also guard against reusing an existing `keyId` that already appears in encrypted cell payloads
        // (for example if the encrypted range metadata was removed but ciphertext remains).
        // Generating new key material for the same id would make those cells undecryptable.
        try {
          const checkCells: Array<{ sheetId: string; row: number; col: number }> = [
            { sheetId, row: range.startRow, col: range.startCol },
          ];
          try {
            const getActiveCell = (app as any).getActiveCell;
            if (typeof getActiveCell === "function") {
              const active = getActiveCell.call(app);
              if (active && Number.isInteger(active.row) && active.row >= 0 && Number.isInteger(active.col) && active.col >= 0) {
                checkCells.push({ sheetId, row: active.row, col: active.col });
              }
            }
          } catch {
            // ignore
          }

          for (const cell of checkCells) {
            const usedKeyId = keyIdFromEncryptedCellPayload(session as any, cell);
            if (usedKeyId && usedKeyId === keyId) {
              showToast(
                `Key ID "${keyId}" is already used by encrypted cells. Import the key first (or choose a new Key ID) to avoid key conflicts.`,
                "warning",
              );
              return;
            }
          }
        } catch {
          // Best-effort; don't block encryption if we can't inspect existing cell payloads.
        }
      }

      const confirmed = await showQuickPick(
        [
          {
            label: `Encrypt ${sheetName}!${a1}`,
            description: isReusingKey
              ? `Key ID: ${displayKeyId} (reuse existing key)`
              : existingKeyLookupFailed
                ? `Key ID: ${displayKeyId} (could not verify existing key; will create new key)`
                : `Key ID: ${displayKeyId} (new key)`,
            value: "encrypt",
          },
        ],
        { placeHolder: "Confirm encryption" },
      );
      if (!confirmed) return;

      let keyBytes: Uint8Array;
      let storedKeyId = displayKeyId;
      let didStoreNewKey = false;
      const canSafelyDeleteStoredKeyOnFailure = !existingKeyLookupFailed;
      if (existingKeyBytes) {
        keyBytes = existingKeyBytes;
        storedKeyId = displayKeyId;
      } else {
        keyBytes = randomKeyBytes();
        try {
          const result = await keyStore.set(docId, keyId, bytesToBase64(keyBytes));
          storedKeyId = result.keyId;
          didStoreNewKey = true;
        } catch {
          showToast("Failed to store encryption key.", "error");
          return;
        }
      }

      const createdBy = session.getPermissions()?.userId ?? undefined;
      try {
        manager.add({ sheetId, ...range, keyId: storedKeyId, createdAt: Date.now(), ...(createdBy ? { createdBy } : {}) });
      } catch (err) {
        // If we generated and stored a brand-new key id, clean it up on failure to avoid leaving orphaned keys.
        // (If we could not verify whether the key already existed, do not delete.)
        let deletedStoredKey = false;
        if (didStoreNewKey && canSafelyDeleteStoredKeyOnFailure) {
          try {
            const deleteKey = (keyStore as any)?.delete;
            if (typeof deleteKey === "function") {
              await deleteKey.call(keyStore, docId, storedKeyId);
              deletedStoredKey = true;
            }
          } catch {
            // Best-effort; ignore delete failures.
          }
        }
        const message = err instanceof Error ? err.message : String(err);
        const orphanedNote = didStoreNewKey && !deletedStoredKey ? " (note: key may have been stored locally)" : "";
        showToast(`Failed to encrypt range: ${message}${orphanedNote}`, "error");
        return;
      }

      const exportString = serializeEncryptionKeyExportString({ docId, keyId: storedKeyId, keyBytes });
      void tryCopyToClipboard(exportString);
      showToast(`Encrypted ${sheetName}!${a1}\n${exportString}`, "info", { timeoutMs: 10_000 });
    },
    {
      category: COMMAND_CATEGORY,
      description: "Encrypt the current selection and generate a shareable key.",
      keywords: ["encrypt", "encryption", "protected range", "collaboration", "collaboration:"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "collab.removeEncryptedRange",
    "Remove encrypted range…",
    async () => {
      const session = app.getCollabSession();
      if (!session) {
        showToast("This command requires collaboration mode.", "warning");
        return;
      }
      const role = session.getRole();
      if (!roleCanEncrypt(role)) {
        showToast("You must have an editor role to remove encrypted ranges.", "warning");
        return;
      }

      const manager = getEncryptionManager(app);
      if (!manager) {
        showToast("Encryption is not available for this workbook.", "error");
        return;
      }

      const ranges = app.getSelectionRanges();
      if (!ranges.length) {
        showToast("Select a range to remove encryption from.", "warning");
        return;
      }
      if (ranges.length > 1) {
        showToast("Remove encrypted range currently supports a single rectangular selection.", "warning");
        return;
      }

      const sheetId = app.getCurrentSheetId();
      const sheetName = app.getCurrentSheetDisplayName();
      const selection = normalizeRange(ranges[0]!);

      const resolveSheetNameById = (id: string): string | null => {
        try {
          const name = app.getSheetDisplayNameById(id);
          return typeof name === "string" ? name : null;
        } catch {
          return null;
        }
      };

      const matchesSheet = (rangeSheetId: string): boolean => {
        const rangeId = String(rangeSheetId ?? "").trim();
        if (!rangeId) return false;
        if (rangeId === sheetId) return true;
        if (rangeId.toLowerCase() === sheetId.toLowerCase()) return true;

        const rangeName = resolveSheetNameById(rangeId);
        // Avoid sheet id/name ambiguity: if the range sheet reference is a valid stable sheet id
        // (i.e. we can resolve it to a display name different from the id) and it doesn't match
        // the current sheet id, do not treat it as a sheet *name*.
        if (rangeName && rangeName !== rangeId) return false;

        return normalizeSheetNameForCompare(rangeId) === normalizeSheetNameForCompare(sheetName);
      };

      let allRanges: ReturnType<typeof manager.list>;
      try {
        allRanges = manager.list();
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        showToast(`Failed to read encrypted ranges: ${message}`, "error");
        return;
      }

      const candidates = allRanges
        .filter((r) => matchesSheet(r.sheetId))
        .filter((r) =>
          rangesIntersect(
            { startRow: r.startRow, startCol: r.startCol, endRow: r.endRow, endCol: r.endCol },
            selection,
          ),
        );

      if (candidates.length === 0) {
        showToast(`No encrypted ranges overlap ${sheetName}!${rangeToA1(selection)}.`, "info");
        return;
      }

      const idToRemove =
        candidates.length === 1
          ? candidates[0]!.id
          : await showQuickPick(
              [
                ...(candidates.length > 1
                  ? [
                      {
                        label: "Remove all overlapping encrypted ranges",
                        description: `${candidates.length} range(s)`,
                        value: "__all__",
                      } as const,
                    ]
                  : []),
                ...candidates.map((r) => {
                  const a1 = rangeToA1({ startRow: r.startRow, startCol: r.startCol, endRow: r.endRow, endCol: r.endCol });
                  const displaySheetName = app.getSheetDisplayNameById(r.sheetId);
                  return {
                    label: `Remove ${displaySheetName}!${a1}`,
                    description: `Key ID: ${r.keyId}`,
                    value: r.id,
                  };
                }),
              ],
              { placeHolder: "Select encrypted range to remove" },
            );

      if (!idToRemove) return;

      const idsToRemove =
        idToRemove === "__all__" ? candidates.map((c) => c.id) : [idToRemove];
      for (const id of idsToRemove) {
        try {
          manager.remove(id);
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          showToast(`Failed to remove encrypted range: ${message}`, "error");
          return;
        }
      }

      showToast(
        idsToRemove.length === 1
          ? "Encrypted range removed (existing encrypted cells remain encrypted)."
          : `Removed ${idsToRemove.length} encrypted ranges (existing encrypted cells remain encrypted).`,
        "info",
      );
    },
    {
      category: COMMAND_CATEGORY,
      description: "Remove encrypted range metadata overlapping the current selection (does not decrypt existing cells).",
      keywords: ["encrypt", "encryption", "remove", "protected range", "collaboration", "collaboration:"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "collab.listEncryptedRanges",
    "List encrypted ranges…",
    async () => {
      const session = app.getCollabSession();
      if (!session) {
        showToast("This command requires collaboration mode.", "warning");
        return;
      }

      const manager = getEncryptionManager(app);
      if (!manager) {
        showToast("Encryption is not available for this workbook.", "error");
        return;
      }

      let ranges: ReturnType<typeof manager.list>;
      try {
        ranges = manager.list();
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        showToast(`Failed to read encrypted ranges: ${message}`, "error");
        return;
      }
      if (ranges.length === 0) {
        showToast("No encrypted ranges in this workbook.", "info");
        return;
      }

      const resolveSheetIdForRange = (rangeSheetId: string): string => {
        const raw = String(rangeSheetId ?? "").trim();
        if (!raw) return raw;

        const getSheetIdByName = (app as any).getSheetIdByName;
        const resolveByName = (name: string): string | null => {
          if (typeof getSheetIdByName !== "function") return null;
          try {
            const resolved = getSheetIdByName.call(app, name);
            const trimmed = typeof resolved === "string" ? resolved.trim() : "";
            return trimmed || null;
          } catch {
            return null;
          }
        };

        // If this looks like a stable sheet id (it resolves to a different display name),
        // prefer it as such and avoid interpreting it as a sheet *name*. For navigation,
        // best-effort resolve the canonical stable id via the display name so we avoid
        // case-mismatched ids in legacy data.
        let displayName: string | null = null;
        try {
          const name = app.getSheetDisplayNameById(raw);
          displayName = typeof name === "string" && name.trim() ? name.trim() : null;
        } catch {
          // ignore
        }
        if (displayName && displayName !== raw) {
          return resolveByName(displayName) ?? raw;
        }

        // Legacy shape: sheet display name instead of stable id.
        return resolveByName(raw) ?? raw;
      };

      const selected = await showQuickPick(
        ranges.map((r) => {
          const a1 = rangeToA1({ startRow: r.startRow, startCol: r.startCol, endRow: r.endRow, endCol: r.endCol });
          const sheetId = resolveSheetIdForRange(r.sheetId);
          const displaySheetName = app.getSheetDisplayNameById(sheetId);
          const label = `${displaySheetName}!${a1}`;
          const description = `Key ID: ${r.keyId}`;
          return { label, description, value: r };
        }),
        { placeHolder: "Select an encrypted range" },
      );
      if (!selected) return;

      const selectRange = (app as any).selectRange;
      if (typeof selectRange !== "function") {
        showToast("Selection API is not available for this workbook.", "error");
        return;
      }

      try {
        selectRange.call(
          app,
          {
            sheetId: resolveSheetIdForRange(selected.sheetId),
            range: {
              startRow: selected.startRow,
              startCol: selected.startCol,
              endRow: selected.endRow,
              endCol: selected.endCol,
            },
          },
          { scrollIntoView: true, focus: true },
        );
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        showToast(`Failed to select encrypted range: ${message}`, "error");
      }
    },
    {
      category: COMMAND_CATEGORY,
      description: "Show all encrypted ranges in the workbook and jump to one.",
      keywords: ["encrypt", "encryption", "list", "protected range", "collaboration", "collaboration:"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "collab.exportEncryptionKey",
    "Export encryption key…",
    async () => {
      const session = app.getCollabSession();
      if (!session) {
        showToast("This command requires collaboration mode.", "warning");
        return;
      }

      const keyStore = app.getCollabEncryptionKeyStore();
      if (!keyStore) {
        showToast("Encryption is not available for this workbook.", "error");
        return;
      }

      const active = app.getActiveCell();
      const sheetId = app.getCurrentSheetId();
      const sheetName = app.getCurrentSheetDisplayName();
      const docId = session.doc.guid;
      const cell = { sheetId, row: active.row, col: active.col };
      // Prefer the key id from an existing encrypted payload (when available locally) so callers
      // can export the correct key even if the encrypted range policy has since changed (e.g.
      // key rotation via add/remove overrides).
      const keyIdFromEnc = keyIdFromEncryptedCellPayload(session as any, cell);
      const policy = createEncryptionPolicyFromDoc(session.doc);
      const keyIdFromPolicy = policy.keyIdForCell(cell);
      if (!keyIdFromEnc && !keyIdFromPolicy) {
        // If the policy fails closed (unknown encryptedRanges schema), shouldEncryptCell returns true
        // for all valid cells but keyIdForCell returns null. In that case, surface a more actionable
        // error rather than incorrectly claiming the cell isn't encrypted.
        if (policy.shouldEncryptCell(cell)) {
          showToast("Encrypted range metadata is in an unsupported format; cannot determine the key id for this cell.", "error");
          return;
        }
        showToast(`The active cell is not inside an encrypted range (${sheetName}).`, "warning");
        return;
      }

      const loadKeyBytes = async (keyId: string): Promise<Uint8Array | null> => {
        const cached = keyStore.getCachedKey(docId, keyId);
        let keyBytes: Uint8Array | null = cached?.keyBytes ?? null;
        if (keyBytes) return keyBytes;
        try {
          const entry = await keyStore.get(docId, keyId);
          if (entry) {
            keyBytes = base64ToBytes(entry.keyBytesBase64);
          }
        } catch {
          // ignore
        }
        return keyBytes;
      };

      // Match the desktop session's key resolution precedence:
      // - If the active cell has an `enc` payload and we have that key locally, export it.
      // - Otherwise, fall back to the policy key id (supports key rotation/overwrite flows).
      const keyIdCandidates = [keyIdFromEnc, keyIdFromPolicy].filter((id): id is string => Boolean(id));

      let keyId: string | null = null;
      let keyBytes: Uint8Array | null = null;
      for (const candidate of keyIdCandidates) {
        const bytes = await loadKeyBytes(candidate);
        if (!bytes) continue;
        keyId = candidate;
        keyBytes = bytes;
        break;
      }

      if (!keyId || !keyBytes) {
        showToast("Missing encryption key for this range. Import the key first.", "warning");
        return;
      }

      const exportString = serializeEncryptionKeyExportString({ docId, keyId, keyBytes });
      const copied = await tryCopyToClipboard(exportString);
      if (copied) showToast("Encryption key copied to clipboard.", "info");

      await showInputBox({
        prompt: "Encryption key (share out-of-band)",
        type: "textarea",
        value: exportString,
        okLabel: "Done",
      });
    },
    {
      category: COMMAND_CATEGORY,
      description: "Export the encryption key for the encrypted range containing the active cell.",
      keywords: ["encrypt", "encryption", "export", "key", "collaboration", "collaboration:"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "collab.importEncryptionKey",
    "Import encryption key…",
    async () => {
      const session = app.getCollabSession();
      if (!session) {
        showToast("This command requires collaboration mode.", "warning");
        return;
      }

      const value = await showInputBox({
        prompt: "Paste encryption key",
        type: "textarea",
        value: "",
        okLabel: "Import",
      });
      if (!value) return;

      let parsed: { docId: string; keyId: string; keyBytes: Uint8Array };
      try {
        parsed = parseEncryptionKeyExportString(value);
      } catch {
        showToast("Invalid encryption key.", "error");
        return;
      }

      const currentDocId = session.doc.guid;
      if (parsed.docId !== currentDocId) {
        showToast("This key is for a different document.", "error");
        return;
      }

      const keyStore = app.getCollabEncryptionKeyStore();
      if (!keyStore) {
        showToast("Encryption is not available for this workbook.", "error");
        return;
      }

      // Guard against accidentally overwriting an existing key id with different bytes
      // (which would make previously-encrypted cells unreadable).
      let existing: Uint8Array | null = null;
      let existingLookupFailed = false;
      try {
        const cached = keyStore.getCachedKey(parsed.docId, parsed.keyId);
        existing = cached?.keyBytes ?? null;
        if (!existing) {
          const entry = await keyStore.get(parsed.docId, parsed.keyId);
          if (entry) existing = base64ToBytes(entry.keyBytesBase64);
        }
      } catch {
        // Best-effort: treat lookup failure as "missing key" but prompt before importing since
        // we cannot guarantee we won't overwrite an existing key id.
        existingLookupFailed = true;
      }

      if (!existing && existingLookupFailed) {
        const proceed = await showQuickPick(
          [
            {
              label: `Import key "${parsed.keyId}"`,
              description: "Could not verify whether this key id already exists. Importing may overwrite an existing key.",
              value: "import",
            },
            { label: "Cancel", value: "cancel" },
          ],
          { placeHolder: "Unable to verify existing key" },
        );
        if (proceed !== "import") return;
      }

      if (existing && !bytesEqual(existing, parsed.keyBytes)) {
        const overwrite = await showQuickPick(
          [
            {
              label: `Overwrite existing key "${parsed.keyId}"`,
              description: "Dangerous: this can make previously-encrypted cells unreadable.",
              value: "overwrite",
            },
            { label: "Cancel", value: "cancel" },
          ],
          { placeHolder: "A different key with this id already exists" },
        );
        if (overwrite !== "overwrite") return;
      }

      const action: "import" | "overwrite" | "already" =
        !existing ? "import" : bytesEqual(existing, parsed.keyBytes) ? "already" : "overwrite";

      let storedKeyId = parsed.keyId;
      if (action !== "already") {
        try {
          const result = await keyStore.set(parsed.docId, parsed.keyId, bytesToBase64(parsed.keyBytes));
          storedKeyId = result.keyId;
        } catch {
          showToast("Failed to store encryption key.", "error");
          return;
        }
      }

      // Refresh the collab binder so any already-encrypted cells are rehydrated with the newly available key.
      try {
        app.rehydrateCollabBinder();
      } catch {
        // Best-effort.
      }
      if (action === "already") {
        showToast(`Encryption key "${storedKeyId}" is already imported.`, "info");
      } else if (action === "overwrite") {
        showToast(`Overwrote encryption key "${storedKeyId}".`, "warning");
      } else {
        showToast(`Imported encryption key "${storedKeyId}".`, "info");
      }
    },
    {
      category: COMMAND_CATEGORY,
      description: "Import an encryption key string shared by another collaborator.",
      keywords: ["encrypt", "encryption", "import", "key", "collaboration", "collaboration:"],
    },
  );
}
