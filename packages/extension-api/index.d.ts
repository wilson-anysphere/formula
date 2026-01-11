export type CellValue = string | number | boolean | null;

export interface Disposable {
  dispose(): void;
}

export interface ExtensionContext {
  readonly extensionId: string;
  readonly extensionPath: string;
  readonly subscriptions: Disposable[];
}

export interface Workbook {
  readonly name: string;
  readonly path?: string | null;
}

export interface Sheet {
  readonly id: string;
  readonly name: string;
}

export interface Range {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
  readonly values: CellValue[][];
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

export namespace workbook {
  function getActiveWorkbook(): Promise<Workbook>;
  function openWorkbook(path: string): Promise<Workbook>;
  function createWorkbook(): Promise<Workbook>;
}

export namespace sheets {
  function getActiveSheet(): Promise<Sheet>;
  function getSheet(name: string): Promise<Sheet | undefined>;
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
}

export namespace clipboard {
  function readText(): Promise<string>;
  function writeText(text: string): Promise<void>;
}

export namespace ui {
  type MessageType = "info" | "warning" | "error";

  interface PanelOptions {
    title: string;
    icon?: string;
    position?: "left" | "right" | "bottom";
  }

  function showMessage(message: string, type?: MessageType): Promise<void>;
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
}

export namespace events {
  function onSelectionChanged(callback: (e: { selection: Range }) => void): Disposable;
  function onCellChanged(
    callback: (e: { row: number; col: number; value: CellValue }) => void
  ): Disposable;
  function onSheetActivated(callback: (e: { sheet: Sheet }) => void): Disposable;
  function onWorkbookOpened(callback: (e: { workbook: Workbook }) => void): Disposable;
  function onBeforeSave(callback: (e: { workbook: Workbook }) => void): Disposable;
  function onViewActivated(callback: (e: { viewId: string }) => void): Disposable;
}

export namespace context {
  const extensionId: string;
  const extensionPath: string;
}

/**
 * Internal APIs used by the extension host worker runtime.
 * These are not part of the public extension authoring surface.
 */
export function __setTransport(transport: { postMessage(message: any): void }): void;
export function __setContext(ctx: { extensionId: string; extensionPath: string }): void;
export function __handleMessage(message: any): void;
