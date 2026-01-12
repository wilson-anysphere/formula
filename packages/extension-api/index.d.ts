export type CellValue = string | number | boolean | null;

/**
 * Menu location ids for `contributes.menus` and `formula.ui.registerContextMenu(...)`.
 *
 * Known/reserved locations are documented in `docs/10-extensibility.md`.
 */
export type MenuId =
  | "cell/context"
  | "row/context"
  | "column/context"
  | "corner/context"
  | (string & {});

export interface Disposable {
  dispose(): void;
}

export interface ExtensionContext {
  readonly extensionId: string;
  readonly extensionPath: string;
  readonly extensionUri: string;
  readonly globalStoragePath: string;
  readonly workspaceStoragePath: string;
  readonly subscriptions: Disposable[];
}

export interface Workbook {
  readonly name: string;
  /**
   * Absolute path for file-backed workbooks; `null` for unsaved workbooks.
   */
  readonly path?: string | null;
  readonly sheets: Sheet[];
  readonly activeSheet: Sheet;
  /**
   * Save the active workbook.
   *
   * In desktop builds this may show a native dialog (e.g. Save As for unsaved workbooks).
   * If the user cancels, the Promise rejects with an Error whose `name` is `"AbortError"`.
   */
  save(): Promise<void>;
  /**
   * Save the workbook to a specific path.
   *
   * If the user cancels, the Promise rejects with an Error whose `name` is `"AbortError"`.
   * The API throws if `path` is empty/whitespace (or otherwise not a usable non-empty string).
   */
  saveAs(path: string): Promise<void>;
  /**
   * Close the current workbook.
   *
   * If the user cancels, the Promise rejects with an Error whose `name` is `"AbortError"`.
   */
  close(): Promise<void>;
}

export interface Sheet {
  readonly id: string;
  readonly name: string;
  getRange(ref: string): Promise<Range>;
  setRange(ref: string, values: CellValue[][]): Promise<void>;
  activate(): Promise<Sheet>;
  rename(name: string): Promise<Sheet>;
}

export interface Range {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
  readonly address: string;
  /** 2D array indexed by [row][col] relative to startRow/startCol. */
  readonly values: CellValue[][];
  /** Formula strings for the range (null when not available). */
  readonly formulas: (string | null)[][];
  /**
   * Indicates the range payload was truncated due to size limits. When true, `values`/`formulas`
   * may be empty to avoid allocating multi-million-cell matrices in memory.
   */
  readonly truncated?: boolean;
}

export interface PanelWebview {
  html: string;
  setHtml(html: string): Promise<void>;
  postMessage(message: any): Promise<void>;
  onDidReceiveMessage(handler: (message: any) => void): Disposable;
}

export interface Panel extends Disposable {
  readonly id: string;
  readonly webview: PanelWebview;
}

/**
 * Workbook lifecycle APIs.
 *
 * Most operations require the `workbook.manage` permission.
 */
export namespace workbook {
  function getActiveWorkbook(): Promise<Workbook>;
  /**
   * Open a workbook from a file path.
   *
   * The API throws if `path` is empty/whitespace (or otherwise not a usable non-empty string). In desktop builds this may prompt the user
   * to discard unsaved changes; cancellations reject with `AbortError`.
   */
  function openWorkbook(path: string): Promise<Workbook>;
  /**
   * Create a new blank workbook.
   *
   * In desktop builds this may prompt the user to discard unsaved changes; cancellations reject
   * with `AbortError`.
   */
  function createWorkbook(): Promise<Workbook>;
  /**
   * Save the active workbook. For unsaved workbooks this may behave like Save As.
   */
  function save(): Promise<void>;
  /**
   * Save the active workbook to a specific path (Save As).
   *
   * The API throws if `path` is empty/whitespace (or otherwise not a usable non-empty string).
   */
  function saveAs(path: string): Promise<void>;
  /**
   * Close the current workbook.
   *
   * In desktop builds this may prompt the user to discard unsaved changes; cancellations reject
   * with `AbortError`.
   */
  function close(): Promise<void>;
}

export namespace sheets {
  function getActiveSheet(): Promise<Sheet>;
  function getSheet(name: string): Promise<Sheet | undefined>;
  function activateSheet(name: string): Promise<Sheet>;
  function createSheet(name: string): Promise<Sheet>;
  function renameSheet(oldName: string, newName: string): Promise<void>;
  function deleteSheet(name: string): Promise<void>;
}

export namespace cells {
  function getSelection(): Promise<Range>;
  function getRange(ref: string): Promise<Range>;
  function getCell(row: number, col: number): Promise<CellValue>;
  function setCell(row: number, col: number, value: CellValue): Promise<void>;
  function setRange(ref: string, values: CellValue[][]): Promise<void>;
}

export namespace commands {
  function registerCommand(
    id: string,
    handler: (...args: any[]) => any | Promise<any>
  ): Promise<Disposable>;
  function executeCommand(id: string, ...args: any[]): Promise<any>;
}

export namespace functions {
  type CustomFunctionHandler = (...args: any[]) => any | Promise<any>;

  interface CustomFunctionDefinition {
    description?: string;
    parameters?: Array<{ name: string; type: string; description?: string }>;
    result?: { type: string };
    isAsync?: boolean;
    returnsArray?: boolean;
    handler: CustomFunctionHandler;
  }

  function register(name: string, def: CustomFunctionDefinition): Promise<Disposable>;
}

export interface DataConnectorQueryResult {
  columns: string[];
  rows: any[][];
}

export interface DataConnectorImplementation {
  browse(config: any, path?: string | null): Promise<any>;
  query(config: any, query: any): Promise<DataConnectorQueryResult>;
  getConnectionConfig?: (...args: any[]) => Promise<any>;
  testConnection?: (...args: any[]) => Promise<any>;
  getQueryBuilder?: (...args: any[]) => Promise<any>;
}

export namespace dataConnectors {
  function register(connectorId: string, impl: DataConnectorImplementation): Promise<Disposable>;
}

export interface FetchResponse {
  readonly ok: boolean;
  readonly status: number;
  readonly statusText: string;
  readonly url: string;
  readonly headers: {
    get(name: string): string | undefined;
  };
  text(): Promise<string>;
  json<T = any>(): Promise<T>;
}

export namespace network {
  function fetch(url: string, init?: any): Promise<FetchResponse>;
  function openWebSocket(url: string): Promise<void>;
}

export namespace clipboard {
  function readText(): Promise<string>;
  function writeText(text: string): Promise<void>;
}

export namespace ui {
  type MessageType = "info" | "warning" | "error";

  interface InputBoxOptions {
    prompt?: string;
    value?: string;
    placeHolder?: string;
    type?: "text" | "password" | "textarea";
    rows?: number;
    okLabel?: string;
    cancelLabel?: string;
  }

  interface QuickPickItem<T = any> {
    label: string;
    value: T;
    description?: string;
    detail?: string;
  }

  interface QuickPickOptions {
    placeHolder?: string;
  }

  interface MenuItem {
    command: string;
    when?: string;
    group?: string;
  }

  interface PanelOptions {
    title: string;
    icon?: string;
    position?: "left" | "right" | "bottom";
  }

  function showMessage(message: string, type?: MessageType): Promise<void>;
  function showInputBox(options?: InputBoxOptions): Promise<string | undefined>;
  function showQuickPick<T>(items: QuickPickItem<T>[], options?: QuickPickOptions): Promise<T | undefined>;
  function registerContextMenu(menuId: MenuId, items: MenuItem[]): Promise<Disposable>;
  function createPanel(id: string, options: PanelOptions): Promise<Panel>;
}

export interface StorageApi {
  get<T = unknown>(key: string): Promise<T | undefined>;
  set<T = unknown>(key: string, value: T): Promise<void>;
  delete(key: string): Promise<void>;
}

export const storage: StorageApi;

export namespace config {
  function get<T = unknown>(key: string): Promise<T | undefined>;
  function update(key: string, value: any): Promise<void>;
  function onDidChange(callback: (e: { key: string; value: any }) => void): Disposable;
}

export namespace events {
  /**
   * Fired when the active selection changes.
   *
   * Note: in desktop builds with DLP enabled, receiving selection/cell values via events may
   * participate in clipboard taint tracking.
   */
  function onSelectionChanged(callback: (e: { sheetId?: string; selection: Range }) => void): Disposable;
  /**
   * Fired when a cell value changes.
   *
   * Note: in desktop builds with DLP enabled, receiving selection/cell values via events may
   * participate in clipboard taint tracking.
   */
  function onCellChanged(
    callback: (e: { sheetId?: string; row: number; col: number; value: CellValue }) => void
  ): Disposable;
  function onSheetActivated(callback: (e: { sheet: Sheet }) => void): Disposable;
  /**
   * Fired after a workbook is opened/created/closed.
   *
   * In desktop builds this is emitted by both UI and extension-initiated workbook operations.
   */
  function onWorkbookOpened(callback: (e: { workbook: Workbook }) => void): Disposable;
  /**
   * Fired before a workbook save occurs.
   *
   * In desktop builds, saving an unsaved workbook may prompt for a Save As path. In that case
   * the event is emitted only once the path is chosen (so cancelling the dialog does not fire).
   */
  function onBeforeSave(callback: (e: { workbook: Workbook }) => void): Disposable;
  function onViewActivated(callback: (e: { viewId: string }) => void): Disposable;
}

export namespace context {
  const extensionId: string;
  const extensionPath: string;
  const extensionUri: string;
  const globalStoragePath: string;
  const workspaceStoragePath: string;
}

/**
 * Internal APIs used by the extension host worker runtime.
 * These are not part of the public extension authoring surface.
 */
export function __setTransport(transport: { postMessage(message: any): void }): void;
export function __setContext(ctx: {
  extensionId: string;
  extensionPath: string;
  extensionUri?: string;
  globalStoragePath?: string;
  workspaceStoragePath?: string;
}): void;
export function __handleMessage(message: any): void;
