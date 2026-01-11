# Real-Time Collaboration

## Overview

Collaboration must be seamless, conflict-free, and work offline. We use **CRDTs (Conflict-free Replicated Data Types)** via Yjs for the core sync engine, providing better offline support and conflict resolution than traditional Operational Transformation.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  CLIENT A                              CLIENT B                              │
│  ┌────────────┐                       ┌────────────┐                        │
│  │ Local Doc  │                       │ Local Doc  │                        │
│  │ (Yjs)      │                       │ (Yjs)      │                        │
│  └─────┬──────┘                       └─────┬──────┘                        │
│        │                                    │                               │
│        │ Updates                            │ Updates                       │
│        ▼                                    ▼                               │
│  ┌────────────┐                       ┌────────────┐                        │
│  │ Awareness  │                       │ Awareness  │                        │
│  │ Protocol   │                       │ Protocol   │                        │
│  └─────┬──────┘                       └─────┬──────┘                        │
│        │                                    │                               │
│        └──────────────┬─────────────────────┘                               │
│                       ▼                                                     │
│              ┌────────────────┐                                             │
│              │  WebSocket/    │                                             │
│              │  WebRTC Sync   │                                             │
│              └────────┬───────┘                                             │
│                       │                                                     │
│                       ▼                                                     │
│              ┌────────────────┐                                             │
│              │  Sync Server   │                                             │
│              │  (y-websocket) │                                             │
│              └────────┬───────┘                                             │
│                       │                                                     │
│                       ▼                                                     │
│              ┌────────────────┐                                             │
│              │  Persistence   │                                             │
│              │  (Database)    │                                             │
│              └────────────────┘                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## CRDT Data Model

### Yjs Document Structure

```typescript
import * as Y from "yjs";

interface SpreadsheetDoc {
  // Root Y.Doc
  doc: Y.Doc;
  
  // Sheets array
  sheets: Y.Array<Y.Map<any>>;
  
  // Per-sheet cell data
  // Key: "sheetId:row:col", Value: cell data
  cells: Y.Map<Y.Map<any>>;
  
  // Metadata
  metadata: Y.Map<any>;
  
  // Named ranges
  namedRanges: Y.Map<any>;
}

class CollaborativeDocument {
  private doc: Y.Doc;
  
  constructor(documentId: string) {
    this.doc = new Y.Doc({ guid: documentId });
    this.initializeStructure();
  }
  
  private initializeStructure(): void {
    // Get or create top-level structures
    const sheets = this.doc.getArray("sheets");
    const cells = this.doc.getMap("cells");
    const metadata = this.doc.getMap("metadata");
    
    // Initialize default sheet if empty
    if (sheets.length === 0) {
      sheets.push([this.createSheet("Sheet1")]);
    }
  }
  
  private createSheet(name: string): Y.Map<any> {
    const sheet = new Y.Map();
    sheet.set("id", crypto.randomUUID());
    sheet.set("name", name);
    sheet.set("frozenRows", 0);
    sheet.set("frozenCols", 0);
    return sheet;
  }
  
  setCell(sheetId: string, row: number, col: number, value: CellValue, formula?: string): void {
    this.doc.transact(() => {
      const cells = this.doc.getMap("cells");
      const cellKey = `${sheetId}:${row}:${col}`;
      
      let cellData = cells.get(cellKey) as Y.Map<any>;
      if (!cellData) {
        cellData = new Y.Map();
        cells.set(cellKey, cellData);
      }
      
      cellData.set("value", value);
      if (formula) {
        cellData.set("formula", formula);
      } else {
        cellData.delete("formula");
      }
      cellData.set("modified", Date.now());
      cellData.set("modifiedBy", this.userId);
    });
  }
  
  getCell(sheetId: string, row: number, col: number): CellData | null {
    const cells = this.doc.getMap("cells");
    const cellKey = `${sheetId}:${row}:${col}`;
    const cellData = cells.get(cellKey) as Y.Map<any>;
    
    if (!cellData) return null;
    
    return {
      value: cellData.get("value"),
      formula: cellData.get("formula"),
      modified: cellData.get("modified"),
      modifiedBy: cellData.get("modifiedBy")
    };
  }
}
```

### Handling Formulas in CRDT

Formulas require special handling because they're interconnected:

```typescript
class FormulaCollaboration {
  private formulaEngine: FormulaEngine;
  private doc: CollaborativeDocument;
  
  onCellChange(sheetId: string, row: number, col: number, newFormula: string): void {
    // Parse and validate formula locally first
    const parsed = this.formulaEngine.parse(newFormula);
    if (parsed.error) {
      // Don't sync invalid formulas
      return;
    }
    
    // Update CRDT
    this.doc.setCell(sheetId, row, col, null, newFormula);
    
    // Note: We don't sync calculated values
    // Each client recalculates independently for consistency
  }
  
  onRemoteChange(changes: CellChange[]): void {
    for (const change of changes) {
      // Mark cell dirty in dependency graph
      this.formulaEngine.markDirty(change.cell);
    }
    
    // Trigger recalculation
    this.formulaEngine.recalculate();
  }
}
```

---

## Presence Awareness

### Awareness Protocol

```typescript
import { Awareness } from "y-protocols/awareness";

interface UserPresence {
  id: string;
  name: string;
  color: string;
  cursor?: CursorPosition;
  selection?: Selection;
  activeSheet: string;
  lastActive: number;
}

interface CursorPosition {
  row: number;
  col: number;
}

class PresenceManager {
  private awareness: Awareness;
  private localUser: UserPresence;
  
  constructor(doc: Y.Doc, user: UserInfo) {
    this.awareness = new Awareness(doc);
    this.localUser = {
      id: user.id,
      name: user.name,
      color: this.assignColor(user.id),
      activeSheet: "",
      lastActive: Date.now()
    };
    
    // Set local state
    this.awareness.setLocalState(this.localUser);
    
    // Listen for remote changes
    this.awareness.on("change", this.onAwarenessChange.bind(this));
  }
  
  updateCursor(row: number, col: number): void {
    this.localUser.cursor = { row, col };
    this.localUser.lastActive = Date.now();
    this.awareness.setLocalState(this.localUser);
  }
  
  updateSelection(selection: Selection): void {
    this.localUser.selection = selection;
    this.localUser.lastActive = Date.now();
    this.awareness.setLocalState(this.localUser);
  }
  
  getOtherUsers(): UserPresence[] {
    const states = this.awareness.getStates();
    const users: UserPresence[] = [];
    
    states.forEach((state, clientId) => {
      if (clientId !== this.awareness.clientID && state) {
        users.push(state as UserPresence);
      }
    });
    
    return users;
  }
  
  private assignColor(userId: string): string {
    // Generate consistent color from user ID
    const colors = [
      "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4",
      "#FFEAA7", "#DDA0DD", "#98D8C8", "#F7DC6F"
    ];
    const hash = userId.split("").reduce((a, b) => {
      a = ((a << 5) - a) + b.charCodeAt(0);
      return a & a;
    }, 0);
    return colors[Math.abs(hash) % colors.length];
  }
  
  private onAwarenessChange(changes: { added: number[], updated: number[], removed: number[] }): void {
    // Notify UI to update presence indicators
    this.emit("presenceChanged", {
      users: this.getOtherUsers()
    });
  }
}
```

### Rendering Presence

```typescript
class PresenceRenderer {
  render(ctx: CanvasRenderingContext2D, users: UserPresence[], viewport: Viewport): void {
    for (const user of users) {
      // Skip users on different sheets
      if (user.activeSheet !== this.currentSheet) continue;
      
      // Render cursor
      if (user.cursor) {
        this.renderCursor(ctx, user, viewport);
      }
      
      // Render selection
      if (user.selection) {
        this.renderSelection(ctx, user, viewport);
      }
    }
  }
  
  private renderCursor(ctx: CanvasRenderingContext2D, user: UserPresence, viewport: Viewport): void {
    const { row, col } = user.cursor!;
    const bounds = this.getCellBounds(row, col, viewport);
    
    if (!bounds) return;  // Outside viewport
    
    // Draw cursor highlight
    ctx.strokeStyle = user.color;
    ctx.lineWidth = 2;
    ctx.strokeRect(bounds.x, bounds.y, bounds.width, bounds.height);
    
    // Draw name badge
    ctx.fillStyle = user.color;
    ctx.fillRect(bounds.x, bounds.y - 20, this.measureText(user.name) + 8, 18);
    
    ctx.fillStyle = "#FFFFFF";
    ctx.font = "12px sans-serif";
    ctx.fillText(user.name, bounds.x + 4, bounds.y - 6);
  }
  
  private renderSelection(ctx: CanvasRenderingContext2D, user: UserPresence, viewport: Viewport): void {
    const selection = user.selection!;
    
    // Semi-transparent fill
    ctx.fillStyle = user.color + "20";  // 12.5% opacity
    
    for (const range of selection.ranges) {
      const bounds = this.getRangeBounds(range, viewport);
      if (bounds) {
        ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
        
        // Border
        ctx.strokeStyle = user.color + "80";  // 50% opacity
        ctx.lineWidth = 1;
        ctx.strokeRect(bounds.x, bounds.y, bounds.width, bounds.height);
      }
    }
  }
}
```

---

## Sync Protocol

### WebSocket Connection

```typescript
import { WebsocketProvider } from "y-websocket";

class SyncManager {
  private provider: WebsocketProvider;
  private doc: Y.Doc;
  
  constructor(doc: Y.Doc, documentId: string, serverUrl: string) {
    this.doc = doc;
    this.provider = new WebsocketProvider(serverUrl, documentId, doc, {
      connect: true,
      awareness: new Awareness(doc),
      params: {
        // Auth token for connection
        token: this.getAuthToken()
      }
    });
    
    this.setupEventHandlers();
  }
  
  private setupEventHandlers(): void {
    this.provider.on("status", (event: { status: string }) => {
      console.log("Connection status:", event.status);
      this.emit("connectionStatus", event.status);
    });
    
    this.provider.on("sync", (isSynced: boolean) => {
      console.log("Sync status:", isSynced);
      this.emit("syncStatus", isSynced);
    });
    
    // Handle disconnection
    this.provider.on("connection-close", () => {
      this.scheduleReconnect();
    });
  }
  
  private scheduleReconnect(): void {
    // Exponential backoff
    const delay = Math.min(1000 * Math.pow(2, this.reconnectAttempts), 30000);
    setTimeout(() => {
      this.provider.connect();
      this.reconnectAttempts++;
    }, delay);
  }
  
  disconnect(): void {
    this.provider.disconnect();
  }
}
```

### Offline Support

```typescript
import { IndexeddbPersistence } from "y-indexeddb";

class OfflineManager {
  private persistence: IndexeddbPersistence;
  private syncManager: SyncManager;
  
  constructor(doc: Y.Doc, documentId: string) {
    // Persist to IndexedDB
    this.persistence = new IndexeddbPersistence(documentId, doc);
    
    this.persistence.on("synced", () => {
      console.log("Loaded from IndexedDB");
    });
  }
  
  async getOfflineChanges(): Promise<Uint8Array> {
    // Get local state that hasn't been synced
    return Y.encodeStateAsUpdate(this.doc);
  }
  
  async clearOfflineData(): Promise<void> {
    await this.persistence.clearData();
  }
}

class OfflineIndicator {
  private isOnline: boolean = navigator.onLine;
  private pendingChanges: number = 0;
  
  constructor(private syncManager: SyncManager) {
    window.addEventListener("online", () => {
      this.isOnline = true;
      this.updateUI();
    });
    
    window.addEventListener("offline", () => {
      this.isOnline = false;
      this.updateUI();
    });
    
    syncManager.on("syncStatus", (synced: boolean) => {
      if (synced) {
        this.pendingChanges = 0;
      }
      this.updateUI();
    });
  }
  
  incrementPending(): void {
    this.pendingChanges++;
    this.updateUI();
  }
  
  private updateUI(): void {
    if (!this.isOnline) {
      this.showStatus("Offline - changes saved locally");
    } else if (this.pendingChanges > 0) {
      this.showStatus(`Syncing ${this.pendingChanges} changes...`);
    } else {
      this.showStatus("All changes saved");
    }
  }
}
```

---

## Version History

### Change Tracking

```typescript
interface Version {
  id: string;
  timestamp: Date;
  userId: string;
  userName: string;
  description?: string;
  snapshot: Uint8Array;  // Y.Doc state
  changes: ChangeEntry[];
}

interface ChangeEntry {
  type: "cell" | "row" | "column" | "sheet" | "format";
  target: string;
  before: any;
  after: any;
}

class VersionManager {
  private versions: Version[] = [];
  private doc: Y.Doc;
  private autoSaveInterval: number = 5 * 60 * 1000;  // 5 minutes
  
  constructor(doc: Y.Doc) {
    this.doc = doc;
    this.startAutoSave();
  }
  
  createVersion(description?: string): Version {
    const snapshot = Y.encodeStateAsUpdate(this.doc);
    const changes = this.getChangesSinceLastVersion();
    
    const version: Version = {
      id: crypto.randomUUID(),
      timestamp: new Date(),
      userId: this.currentUserId,
      userName: this.currentUserName,
      description,
      snapshot,
      changes
    };
    
    this.versions.push(version);
    this.persistVersion(version);
    
    return version;
  }
  
  async restoreVersion(versionId: string): Promise<void> {
    const version = this.versions.find(v => v.id === versionId);
    if (!version) throw new Error("Version not found");
    
    // Create new doc from snapshot
    const restoredDoc = new Y.Doc();
    Y.applyUpdate(restoredDoc, version.snapshot);
    
    // Apply to current doc (this will sync to all clients)
    this.doc.transact(() => {
      // Clear current state
      const cells = this.doc.getMap("cells");
      cells.forEach((_, key) => cells.delete(key));
      
      // Apply restored state
      const restoredCells = restoredDoc.getMap("cells");
      restoredCells.forEach((value, key) => {
        cells.set(key, value);
      });
    });
  }
  
  private startAutoSave(): void {
    setInterval(() => {
      if (this.hasChangesSinceLastVersion()) {
        this.createVersion("Auto-save");
      }
    }, this.autoSaveInterval);
  }
}
```

### Named Checkpoints

```typescript
interface Checkpoint extends Version {
  name: string;
  isLocked: boolean;
  annotations?: string;
}

class CheckpointManager {
  async createCheckpoint(name: string, annotations?: string): Promise<Checkpoint> {
    const version = this.versionManager.createVersion(name);
    
    const checkpoint: Checkpoint = {
      ...version,
      name,
      isLocked: false,
      annotations
    };
    
    await this.saveCheckpoint(checkpoint);
    return checkpoint;
  }
  
  async lockCheckpoint(checkpointId: string): Promise<void> {
    const checkpoint = await this.getCheckpoint(checkpointId);
    checkpoint.isLocked = true;
    await this.saveCheckpoint(checkpoint);
  }
  
  async listCheckpoints(): Promise<Checkpoint[]> {
    return this.db.query("checkpoints", {
      orderBy: "timestamp",
      order: "desc"
    });
  }
}
```

---

## Conflict Resolution

### Cell-Level Conflicts

CRDTs handle most conflicts automatically via last-writer-wins. For cases where we need smarter resolution:

```typescript
interface ConflictResolution {
  strategy: "last_write_wins" | "first_write_wins" | "merge" | "prompt_user";
  mergeFunction?: (a: CellValue, b: CellValue) => CellValue;
}

class ConflictHandler {
  // For most cells, last-write-wins is fine
  defaultResolution: ConflictResolution = { strategy: "last_write_wins" };
  
  // For specific cells (e.g., counters, totals), we might want merge
  specialResolutions: Map<string, ConflictResolution> = new Map();
  
  registerMergeCell(cellRef: string, mergeFunction: (a: any, b: any) => any): void {
    this.specialResolutions.set(cellRef, {
      strategy: "merge",
      mergeFunction
    });
  }
  
  // Example: Counter cell that should sum concurrent increments
  setupCounterCell(cellRef: string): void {
    this.registerMergeCell(cellRef, (a: number, b: number) => {
      // This is a simplified example; real implementation would track deltas
      return a + b;
    });
  }
}
```

### Formula Conflicts

When formulas conflict (multiple users edit same cell's formula):

```typescript
class FormulaConflictResolver {
  async handleFormulaConflict(
    cell: CellRef,
    localFormula: string,
    remoteFormula: string,
    remoteUser: string
  ): Promise<string> {
    // If formulas are equivalent, no conflict
    if (this.formulasEquivalent(localFormula, remoteFormula)) {
      return remoteFormula;  // Use remote (it's equivalent anyway)
    }
    
    // Check if one is a subset/extension of other
    if (this.isExtension(localFormula, remoteFormula)) {
      return remoteFormula;
    }
    if (this.isExtension(remoteFormula, localFormula)) {
      return localFormula;
    }
    
    // True conflict - prompt user
    const resolution = await this.promptUser({
      cell,
      localFormula,
      remoteFormula,
      remoteUser,
      options: [
        { label: "Keep yours", value: localFormula },
        { label: `Use ${remoteUser}'s`, value: remoteFormula },
        { label: "Edit manually", value: "edit" }
      ]
    });
    
    return resolution;
  }
  
  private formulasEquivalent(a: string, b: string): boolean {
    // Normalize and compare ASTs
    const astA = this.parse(a);
    const astB = this.parse(b);
    return this.astEqual(astA, astB);
  }
}
```

---

## Semantic Diff

### Cell-by-Cell Comparison

```typescript
interface DiffResult {
  added: CellChange[];
  removed: CellChange[];
  modified: CellChange[];
  moved: MoveChange[];
  formatOnly: CellChange[];
}

interface CellChange {
  cell: CellRef;
  oldValue?: CellValue;
  newValue?: CellValue;
  oldFormula?: string;
  newFormula?: string;
}

interface MoveChange {
  oldLocation: CellRef;
  newLocation: CellRef;
  value: CellValue;
}

class SemanticDiff {
  compare(before: SheetData, after: SheetData): DiffResult {
    const result: DiffResult = {
      added: [],
      removed: [],
      modified: [],
      moved: [],
      formatOnly: []
    };
    
    // Find removed and modified cells
    for (const [cellId, beforeCell] of before.cells) {
      const afterCell = after.cells.get(cellId);
      
      if (!afterCell) {
        // Check if it moved
        const movedTo = this.findMovedCell(beforeCell, after);
        if (movedTo) {
          result.moved.push({
            oldLocation: this.parseCellId(cellId),
            newLocation: movedTo,
            value: beforeCell.value
          });
        } else {
          result.removed.push({
            cell: this.parseCellId(cellId),
            oldValue: beforeCell.value,
            oldFormula: beforeCell.formula
          });
        }
      } else if (!this.cellsEqual(beforeCell, afterCell)) {
        if (this.onlyFormatChanged(beforeCell, afterCell)) {
          result.formatOnly.push({
            cell: this.parseCellId(cellId),
            oldValue: beforeCell.value,
            newValue: afterCell.value
          });
        } else {
          result.modified.push({
            cell: this.parseCellId(cellId),
            oldValue: beforeCell.value,
            newValue: afterCell.value,
            oldFormula: beforeCell.formula,
            newFormula: afterCell.formula
          });
        }
      }
    }
    
    // Find added cells
    for (const [cellId, afterCell] of after.cells) {
      if (!before.cells.has(cellId)) {
        // Check if it was a move (already handled)
        const wasMove = result.moved.some(m => 
          this.cellRefsEqual(m.newLocation, this.parseCellId(cellId))
        );
        
        if (!wasMove) {
          result.added.push({
            cell: this.parseCellId(cellId),
            newValue: afterCell.value,
            newFormula: afterCell.formula
          });
        }
      }
    }
    
    return result;
  }
  
  private findMovedCell(cell: Cell, after: SheetData): CellRef | null {
    // Look for cell with same value and formula in different location
    for (const [cellId, afterCell] of after.cells) {
      if (this.cellsEqual(cell, afterCell)) {
        return this.parseCellId(cellId);
      }
    }
    return null;
  }
}
```

### Diff Visualization

```typescript
class DiffRenderer {
  render(diff: DiffResult, ctx: CanvasRenderingContext2D, viewport: Viewport): void {
    // Added cells - green highlight
    for (const change of diff.added) {
      const bounds = this.getCellBounds(change.cell, viewport);
      if (bounds) {
        ctx.fillStyle = "rgba(0, 200, 0, 0.2)";
        ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
        
        ctx.strokeStyle = "rgba(0, 200, 0, 0.8)";
        ctx.lineWidth = 2;
        ctx.strokeRect(bounds.x, bounds.y, bounds.width, bounds.height);
      }
    }
    
    // Removed cells - red highlight
    for (const change of diff.removed) {
      const bounds = this.getCellBounds(change.cell, viewport);
      if (bounds) {
        ctx.fillStyle = "rgba(200, 0, 0, 0.2)";
        ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
        
        // Strikethrough
        ctx.strokeStyle = "rgba(200, 0, 0, 0.8)";
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(bounds.x, bounds.y + bounds.height / 2);
        ctx.lineTo(bounds.x + bounds.width, bounds.y + bounds.height / 2);
        ctx.stroke();
      }
    }
    
    // Modified cells - yellow highlight
    for (const change of diff.modified) {
      const bounds = this.getCellBounds(change.cell, viewport);
      if (bounds) {
        ctx.fillStyle = "rgba(255, 200, 0, 0.2)";
        ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
        
        ctx.strokeStyle = "rgba(255, 200, 0, 0.8)";
        ctx.lineWidth = 2;
        ctx.strokeRect(bounds.x, bounds.y, bounds.width, bounds.height);
      }
    }
  }
}
```

---

## Comments and Annotations

### Cell-Level Comments

**Canonical schema:** `comments` is a `Y.Map` keyed by comment id (`Y.Map<string, Y.Map>`).

> Compatibility note: older documents stored `comments` as a `Y.Array<Y.Map>` (one entry per
> comment). Because Yjs root types are schema-defined by name, calling `doc.getMap("comments")`
> on an Array-backed root can make the legacy array content inaccessible. When loading unknown
> docs, detect the root kind first (or run a migration) before choosing a constructor
> (e.g. `getCommentsRoot` / `migrateCommentsArrayToMap` in `@formula/collab-comments`; the
> migration renames the legacy array root to `comments_legacy*` and creates the canonical map
> under `comments`).

```typescript
interface Comment {
  id: string;
  cellRef: CellRef;
  author: string;
  authorId: string;
  content: string;
  createdAt: Date;
  resolved: boolean;
  replies: Reply[];
}

interface Reply {
  id: string;
  author: string;
  authorId: string;
  content: string;
  createdAt: Date;
}

class CommentManager {
  // Keyed by comment id
  private comments: Y.Map<Y.Map<any>>;
  
  constructor(doc: Y.Doc) {
    this.comments = doc.getMap("comments");
  }
  
  addComment(cellRef: CellRef, content: string): Comment {
    const comment: Comment = {
      id: crypto.randomUUID(),
      cellRef,
      author: this.currentUserName,
      authorId: this.currentUserId,
      content,
      createdAt: new Date(),
      resolved: false,
      replies: []
    };
    
    const yComment = new Y.Map();
    Object.entries(comment).forEach(([key, value]) => {
      yComment.set(key, value);
    });
    
    this.comments.set(comment.id, yComment);
    return comment;
  }
  
  addReply(commentId: string, content: string): Reply {
    const comment = this.findComment(commentId);
    if (!comment) throw new Error("Comment not found");
    
    const reply: Reply = {
      id: crypto.randomUUID(),
      author: this.currentUserName,
      authorId: this.currentUserId,
      content,
      createdAt: new Date()
    };
    
    const replies = comment.get("replies") as Y.Array<Y.Map<any>>;
    const yReply = new Y.Map();
    Object.entries(reply).forEach(([key, value]) => {
      yReply.set(key, value);
    });
    replies.push([yReply]);
    
    return reply;
  }
  
  resolveComment(commentId: string): void {
    const comment = this.findComment(commentId);
    if (comment) {
      comment.set("resolved", true);
    }
  }
  
  getCommentsForCell(cellRef: CellRef): Comment[] {
    return Array.from(this.comments.values())
      .filter(c => this.cellRefsEqual(c.get("cellRef"), cellRef))
      .map(c => this.yMapToComment(c));
  }
}
```

---

## Permissions

### Access Control Levels

```typescript
type PermissionLevel = "owner" | "editor" | "commenter" | "viewer";

interface DocumentPermission {
  documentId: string;
  userId: string;
  email: string;
  level: PermissionLevel;
  grantedBy: string;
  grantedAt: Date;
  expiresAt?: Date;
}

interface RangePermission {
  rangeId: string;
  range: Range;
  permissions: CellPermission[];
}

interface CellPermission {
  userId: string;
  canView: boolean;
  canEdit: boolean;
}

class PermissionManager {
  canEdit(userId: string, cellRef: CellRef): boolean {
    // Check document-level permission
    const docPerm = this.getDocumentPermission(userId);
    if (docPerm.level === "viewer" || docPerm.level === "commenter") {
      return false;
    }
    
    // Check cell-level restrictions
    const cellPerm = this.getCellPermission(userId, cellRef);
    if (cellPerm && !cellPerm.canEdit) {
      return false;
    }
    
    // Check if cell is locked
    const cell = this.getCell(cellRef);
    if (cell?.protection?.locked && this.sheetProtected) {
      return false;
    }
    
    return true;
  }
  
  async grantAccess(email: string, level: PermissionLevel): Promise<void> {
    const permission: DocumentPermission = {
      documentId: this.documentId,
      userId: await this.resolveUserId(email),
      email,
      level,
      grantedBy: this.currentUserId,
      grantedAt: new Date()
    };
    
    await this.savePermission(permission);
    await this.sendInvitationEmail(email, level);
  }
  
  async restrictRange(range: Range, allowedUsers: string[]): Promise<void> {
    const rangePermission: RangePermission = {
      rangeId: crypto.randomUUID(),
      range,
      permissions: allowedUsers.map(userId => ({
        userId,
        canView: true,
        canEdit: true
      }))
    };
    
    await this.saveRangePermission(rangePermission);
  }
}
```

---

## Performance Considerations

### Optimizing Sync

```typescript
class SyncOptimizer {
  private batchInterval = 50;  // ms
  private pendingUpdates: Y.Transaction[] = [];
  private batchTimeout: number | null = null;
  
  queueUpdate(transaction: Y.Transaction): void {
    this.pendingUpdates.push(transaction);
    
    if (!this.batchTimeout) {
      this.batchTimeout = setTimeout(() => {
        this.flushUpdates();
      }, this.batchInterval);
    }
  }
  
  private flushUpdates(): void {
    if (this.pendingUpdates.length === 0) return;
    
    // Combine all updates into single sync message
    const updates = this.pendingUpdates;
    this.pendingUpdates = [];
    this.batchTimeout = null;
    
    // Send combined update
    this.sendUpdate(Y.mergeUpdates(updates.map(t => t.update)));
  }
}
```

### Reducing Awareness Overhead

```typescript
class ThrottledAwareness {
  private lastCursorUpdate = 0;
  private cursorThrottle = 100;  // ms
  
  updateCursor(row: number, col: number): void {
    const now = Date.now();
    if (now - this.lastCursorUpdate < this.cursorThrottle) {
      return;  // Throttle
    }
    
    this.lastCursorUpdate = now;
    this.awareness.setLocalState({
      ...this.awareness.getLocalState(),
      cursor: { row, col }
    });
  }
}
```

---

## Testing Strategy

### Collaboration Tests

```typescript
describe("Real-time Collaboration", () => {
  it("syncs cell changes between clients", async () => {
    const doc1 = createCollaborativeDoc("test-doc");
    const doc2 = createCollaborativeDoc("test-doc");
    
    await connectBothClients(doc1, doc2);
    
    // Client 1 makes change
    doc1.setCell("sheet1", 0, 0, "Hello");
    
    // Wait for sync
    await waitForSync(doc2);
    
    // Client 2 should see change
    expect(doc2.getCell("sheet1", 0, 0)?.value).toBe("Hello");
  });
  
  it("handles concurrent edits to different cells", async () => {
    const doc1 = createCollaborativeDoc("test-doc");
    const doc2 = createCollaborativeDoc("test-doc");
    
    await connectBothClients(doc1, doc2);
    
    // Concurrent edits
    doc1.setCell("sheet1", 0, 0, "A1 from client 1");
    doc2.setCell("sheet1", 0, 1, "B1 from client 2");
    
    // Wait for sync
    await waitForSync(doc1);
    await waitForSync(doc2);
    
    // Both changes should be present in both docs
    expect(doc1.getCell("sheet1", 0, 0)?.value).toBe("A1 from client 1");
    expect(doc1.getCell("sheet1", 0, 1)?.value).toBe("B1 from client 2");
    expect(doc2.getCell("sheet1", 0, 0)?.value).toBe("A1 from client 1");
    expect(doc2.getCell("sheet1", 0, 1)?.value).toBe("B1 from client 2");
  });
  
  it("handles offline changes and syncs on reconnect", async () => {
    const doc1 = createCollaborativeDoc("test-doc");
    const doc2 = createCollaborativeDoc("test-doc");
    
    await connectBothClients(doc1, doc2);
    
    // Disconnect client 2
    doc2.disconnect();
    
    // Both make changes offline
    doc1.setCell("sheet1", 0, 0, "Online change");
    doc2.setCell("sheet1", 1, 0, "Offline change");
    
    // Reconnect
    await doc2.reconnect();
    await waitForSync(doc1);
    await waitForSync(doc2);
    
    // Both changes should be present
    expect(doc1.getCell("sheet1", 0, 0)?.value).toBe("Online change");
    expect(doc1.getCell("sheet1", 1, 0)?.value).toBe("Offline change");
  });
});
```
