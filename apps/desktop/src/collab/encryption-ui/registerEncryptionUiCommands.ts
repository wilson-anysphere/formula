import type { SpreadsheetApp } from "../../app/spreadsheetApp";
import type { CommandRegistry } from "../../extensions/commandRegistry.js";
import { showInputBox, showQuickPick, showToast } from "../../extensions/ui.js";
import type { Range } from "../../selection/types";
import { rangeToA1 } from "../../selection/a1";

import { base64ToBytes, bytesToBase64 } from "@formula/collab-encryption";
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

      const keyId = await showInputBox({
        prompt: `Encrypt ${sheetName}!${a1} – Key ID`,
        value: randomKeyId(),
        placeHolder: "Key ID (for example: team-budget-q1)",
        okLabel: "Next",
      });
      if (!keyId) return;

      const confirmed = await showQuickPick(
        [
          {
            label: `Encrypt ${sheetName}!${a1}`,
            description: `Key ID: ${keyId}`,
            value: "encrypt",
          },
        ],
        { placeHolder: "Confirm encryption" },
      );
      if (!confirmed) return;

      const keyBytes = randomKeyBytes();
      const docId = session.doc.guid;
      const keyStore = app.getCollabEncryptionKeyStore();
      if (!keyStore) {
        showToast("Encryption is not available for this workbook.", "error");
        return;
      }

      let storedKeyId = keyId;
      try {
        const result = await keyStore.set(docId, keyId, bytesToBase64(keyBytes));
        storedKeyId = result.keyId;
      } catch {
        showToast("Failed to store encryption key.", "error");
        return;
      }

      const createdBy = session.getPermissions()?.userId ?? undefined;
      manager.add({ sheetId, ...range, keyId: storedKeyId, createdAt: Date.now(), ...(createdBy ? { createdBy } : {}) });

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
    "collab.exportEncryptionKey",
    "Export encryption key…",
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

      const active = app.getActiveCell();
      const sheetId = app.getCurrentSheetId();
      const sheetName = app.getCurrentSheetDisplayName();
      const docId = session.doc.guid;
      const policy = createEncryptionPolicyFromDoc(session.doc);
      const keyId = policy.keyIdForCell({ sheetId, row: active.row, col: active.col });
      if (!keyId) {
        showToast(`The active cell is not inside an encrypted range (${sheetName}).`, "warning");
        return;
      }

      const keyStore = app.getCollabEncryptionKeyStore();
      if (!keyStore) {
        showToast("Encryption is not available for this workbook.", "error");
        return;
      }

      const cached = keyStore.getCachedKey(docId, keyId);
      let keyBytes: Uint8Array | null = cached?.keyBytes ?? null;
      if (!keyBytes) {
        try {
          const entry = await keyStore.get(docId, keyId);
          if (entry) {
            keyBytes = base64ToBytes(entry.keyBytesBase64);
          }
        } catch {
          // ignore
        }
      }
      if (!keyBytes) {
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
      try {
        await keyStore.set(parsed.docId, parsed.keyId, bytesToBase64(parsed.keyBytes));
      } catch {
        showToast("Failed to store encryption key.", "error");
        return;
      }

      // Refresh the collab binder so any already-encrypted cells are rehydrated with the newly available key.
      try {
        app.rehydrateCollabBinder();
      } catch {
        // Best-effort.
      }

      showToast(`Imported encryption key "${parsed.keyId}".`, "info");
    },
    {
      category: COMMAND_CATEGORY,
      description: "Import an encryption key string shared by another collaborator.",
      keywords: ["encrypt", "encryption", "import", "key", "collaboration", "collaboration:"],
    },
  );
}
