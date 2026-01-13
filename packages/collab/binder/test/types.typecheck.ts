import * as Y from "yjs";

import {
  bindYjsToDocumentController,
  type BindYjsToDocumentControllerOptions,
  type CellAddress,
  type EncryptionConfig,
  type PermissionsResolver,
  type RoleBasedPermissions,
  type SheetViewState,
} from "../index.js";

// Basic call shape + return type
{
  const ydoc = new Y.Doc();
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController: {} as any,
  });
  void binder.whenIdle();
  binder.destroy();
}

// Options interface is usable standalone
{
  const ydoc = new Y.Doc();

  const encryption: EncryptionConfig = {
    keyForCell: (_cell) => ({ keyId: "k1", keyBytes: new Uint8Array([1, 2, 3]) }),
    shouldEncryptCell: () => true,
    encryptFormat: true,
  };

  const permissionsFn: PermissionsResolver = (cell: CellAddress) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) return { canRead: true, canEdit: true };
    return { canRead: false };
  };

  const roleBased: RoleBasedPermissions = {
    role: "editor",
    restrictions: [
      { sheetId: "Sheet1", startRow: 0, endRow: 10, startCol: 0, endCol: 0, editAllowlist: ["me"] },
    ],
    userId: "me",
  };

  const opts: BindYjsToDocumentControllerOptions = {
    ydoc,
    documentController: {} as any,
    defaultSheetId: "Sheet1",
    userId: "me",
    undoService: {
      transact(fn) {
        fn();
      },
      origin: { type: "test-origin" },
      localOrigins: new Set(),
    },
    encryption,
    permissions: roleBased,
    // Ensure union accepts function resolver too.
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    canReadCell: (_cell) => true,
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    canEditCell: (_cell) => true,
    maskCellValue: (value, _cell) => value,
    canWriteSharedState: () => true,
    maskCellFormat: true,
    onEditRejected: (_deltas) => {},
    formulaConflictsMode: "formula+value",
  };

  bindYjsToDocumentController(opts);
  bindYjsToDocumentController({ ...opts, permissions: permissionsFn });
}

// SheetViewState shape
{
  const view: SheetViewState = {
    frozenRows: 1,
    frozenCols: 2,
    colWidths: { "0": 120 },
    rowHeights: { "0": 24 },
  };
  void view;
}

// Type errors we want to catch
{
  const ydoc = new Y.Doc();

  // @ts-expect-error - requires ydoc
  bindYjsToDocumentController({ documentController: {} as any });

  // @ts-expect-error - requires documentController
  bindYjsToDocumentController({ ydoc });

  // @ts-expect-error - invalid role
  const badRole: RoleBasedPermissions = { role: "superuser" };
  void badRole;

  // @ts-expect-error - keyBytes must be Uint8Array
  const badEnc: EncryptionConfig = { keyForCell: () => ({ keyId: "k", keyBytes: "nope" }) };
  void badEnc;
}
