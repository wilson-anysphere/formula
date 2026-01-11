import "./src/runtime.js";

const api = globalThis[Symbol.for("formula.extensionApi.api")];
if (!api) {
  throw new Error("@formula/extension-api runtime failed to initialize");
}

export const workbook = api.workbook;
export const sheets = api.sheets;
export const cells = api.cells;
export const commands = api.commands;
export const functions = api.functions;
export const dataConnectors = api.dataConnectors;
export const network = api.network;
export const clipboard = api.clipboard;
export const ui = api.ui;
export const storage = api.storage;
export const config = api.config;
export const events = api.events;
export const context = api.context;
export const __setTransport = api.__setTransport;
export const __setContext = api.__setContext;
export const __handleMessage = api.__handleMessage;
