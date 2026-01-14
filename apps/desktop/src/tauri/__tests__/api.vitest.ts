import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";

import {
  getTauriDialogOrNull,
  getTauriDialogOrThrow,
  getTauriDialogConfirmOrNull,
  getTauriDialogMessageOrNull,
  getTauriDialogOpenOrNull,
  getTauriDialogSaveOrNull,
  getTauriEventApiOrNull,
  getTauriEventApiOrThrow,
  getTauriAppGetNameOrNull,
  getTauriAppGetVersionOrNull,
  getTauriInvokeOrNull,
  getTauriInvokeOrThrow,
  hasTauri,
  hasTauriInvoke,
  hasTauriWindowApi,
  hasTauriWindowHandleApi,
  getTauriWindowHandleOrNull,
  getTauriWindowHandleOrThrow,
} from "../api";

describe("tauri/api dynamic accessors", () => {
  const originalTauriDescriptor = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");

  const restoreTauri = () => {
    if (originalTauriDescriptor) {
      Object.defineProperty(globalThis, "__TAURI__", originalTauriDescriptor);
      return;
    }
    // If the property did not exist initially, remove it.
    // (If it is non-configurable for some reason, deletion will fail; ignore.)
    try {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (globalThis as any).__TAURI__;
    } catch {
      // ignore
    }
  };

  beforeEach(() => {
    // Ensure a consistent, configurable starting point even if a prior test defined
    // `__TAURI__` with a throwing getter.
    Object.defineProperty(globalThis, "__TAURI__", { configurable: true, writable: true, value: undefined });
  });

  afterEach(() => {
    restoreTauri();
    vi.restoreAllMocks();
  });

  describe("getTauriDialog*", () => {
    it("returns null / throws when __TAURI__ is missing", () => {
      expect(getTauriDialogOrNull()).toBeNull();
      expect(() => getTauriDialogOrThrow()).toThrowError("Tauri dialog API not available");
    });

    it("treats a throwing __TAURI__ getter as \"missing\" (best-effort hardening)", () => {
      Object.defineProperty(globalThis, "__TAURI__", {
        configurable: true,
        get() {
          throw new Error("Blocked __TAURI__ access");
        },
      });

      expect(hasTauri()).toBe(false);
      expect(getTauriDialogOrNull()).toBeNull();
      expect(getTauriEventApiOrNull()).toBeNull();
      expect(getTauriWindowHandleOrNull()).toBeNull();
      expect(hasTauriWindowApi()).toBe(false);
      expect(hasTauriWindowHandleApi()).toBe(false);

      expect(() => getTauriDialogOrThrow()).toThrowError("Tauri dialog API not available");
      expect(() => getTauriEventApiOrThrow()).toThrowError("Tauri event API not available");
      expect(() => getTauriWindowHandleOrThrow()).toThrowError("Tauri window API not available");
    });

    it("treats throwing nested properties (e.g. dialog getter) as unavailable", () => {
      const tauri: any = {};
      Object.defineProperty(tauri, "dialog", {
        configurable: true,
        get() {
          throw new Error("Blocked dialog access");
        },
      });
      (globalThis as any).__TAURI__ = tauri;

      expect(getTauriDialogOpenOrNull()).toBeNull();
      expect(getTauriDialogSaveOrNull()).toBeNull();
      expect(getTauriDialogConfirmOrNull()).toBeNull();
      expect(getTauriDialogMessageOrNull()).toBeNull();
      expect(getTauriDialogOrNull()).toBeNull();
      expect(() => getTauriDialogOrThrow()).toThrowError("Tauri dialog API not available");
    });

    it("detects the dialog API on __TAURI__.dialog (legacy shape) and preserves method binding", async () => {
      const dialogApi: any = {};
      let openThis: any = null;
      let saveThis: any = null;
      const open = vi.fn(async function () {
        openThis = this;
        return null;
      });
      const save = vi.fn(async function () {
        saveThis = this;
        return null;
      });
      dialogApi.open = open;
      dialogApi.save = save;
      (globalThis as any).__TAURI__ = { dialog: dialogApi };

      const api = getTauriDialogOrThrow();
      await api.open();
      await api.save();

      expect(open).toHaveBeenCalledTimes(1);
      expect(save).toHaveBeenCalledTimes(1);
      expect(open.mock.calls[0]?.length).toBe(0);
      expect(save.mock.calls[0]?.length).toBe(0);
      expect(openThis).toBe(dialogApi);
      expect(saveThis).toBe(dialogApi);
    });

    it("detects the dialog API on __TAURI__.plugin.dialog (tauri v2 plugin shape) and preserves method binding", async () => {
      const dialogApi: any = {};
      let openThis: any = null;
      let saveThis: any = null;
      const open = vi.fn(async function (options?: Record<string, unknown>) {
        openThis = this;
        expect(options).toEqual({ multiple: false });
        return null;
      });
      const save = vi.fn(async function () {
        saveThis = this;
        return null;
      });
      dialogApi.open = open;
      dialogApi.save = save;
      (globalThis as any).__TAURI__ = { plugin: { dialog: dialogApi } };

      const api = getTauriDialogOrThrow();
      await api.open({ multiple: false });
      await api.save();

      expect(open).toHaveBeenCalledTimes(1);
      expect(save).toHaveBeenCalledTimes(1);
      expect(openThis).toBe(dialogApi);
      expect(saveThis).toBe(dialogApi);
    });

    it("detects the dialog API on __TAURI__.plugins.dialog (alternate plugin container shape) and preserves method binding", async () => {
      const dialogApi: any = {};
      let openThis: any = null;
      let saveThis: any = null;
      const open = vi.fn(async function () {
        openThis = this;
        return null;
      });
      const save = vi.fn(async function () {
        saveThis = this;
        return null;
      });
      dialogApi.open = open;
      dialogApi.save = save;
      (globalThis as any).__TAURI__ = { plugins: { dialog: dialogApi } };

      const api = getTauriDialogOrThrow();
      await api.open();
      await api.save();

      expect(openThis).toBe(dialogApi);
      expect(saveThis).toBe(dialogApi);
    });

    it("treats partial dialog APIs as unavailable (e.g. open without save)", () => {
      const open = vi.fn();
      (globalThis as any).__TAURI__ = { dialog: { open } };
      expect(getTauriDialogOrNull()).toBeNull();
      expect(() => getTauriDialogOrThrow()).toThrowError("Tauri dialog API not available");
    });

    it("exposes open/save-only helpers that do not require the full API surface", async () => {
      const dialogApi: any = {};
      let openThis: any = null;
      const open = vi.fn(async function () {
        openThis = this;
        return null;
      });
      dialogApi.open = open;
      (globalThis as any).__TAURI__ = { dialog: dialogApi };

      const openFn = getTauriDialogOpenOrNull();
      expect(openFn).not.toBeNull();
      await openFn!();
      expect(openThis).toBe(dialogApi);
      expect(getTauriDialogSaveOrNull()).toBeNull();
    });

    it("detects confirm() when available and preserves method binding", async () => {
      const dialogApi: any = {};
      let confirmThis: any = null;
      const confirm = vi.fn(async function (message: string, options?: Record<string, unknown>) {
        confirmThis = this;
        expect(message).toBe("Hello");
        expect(options).toEqual({ title: "Formula" });
        return true;
      });
      dialogApi.confirm = confirm;
      (globalThis as any).__TAURI__ = { dialog: dialogApi };

      const confirmFn = getTauriDialogConfirmOrNull();
      expect(confirmFn).not.toBeNull();
      await expect(confirmFn!("Hello", { title: "Formula" })).resolves.toBe(true);
      expect(confirmThis).toBe(dialogApi);
    });

    it("detects message() (or alert()) when available and preserves method binding", async () => {
      const dialogApi: any = {};
      let messageThis: any = null;
      const message = vi.fn(async function (msg: string) {
        messageThis = this;
        expect(msg).toBe("Hi");
      });
      dialogApi.message = message;
      (globalThis as any).__TAURI__ = { dialog: dialogApi };

      const messageFn = getTauriDialogMessageOrNull();
      expect(messageFn).not.toBeNull();
      await messageFn!("Hi");
      expect(messageThis).toBe(dialogApi);
      expect(message.mock.calls[0]?.length).toBe(1);

      const dialogApi2: any = {};
      let alertThis: any = null;
      const alert = vi.fn(async function (msg: string) {
        alertThis = this;
        expect(msg).toBe("Hi");
      });
      dialogApi2.alert = alert;
      (globalThis as any).__TAURI__ = { dialog: dialogApi2 };

      const alertFn = getTauriDialogMessageOrNull();
      expect(alertFn).not.toBeNull();
      await alertFn!("Hi");
      expect(alertThis).toBe(dialogApi2);
      expect(alert.mock.calls[0]?.length).toBe(1);
    });

    it("detects confirm/message under plugin container shapes too", async () => {
      const dialogApi: any = {};
      let confirmThis: any = null;
      let messageThis: any = null;
      const confirm = vi.fn(async function () {
        confirmThis = this;
        return true;
      });
      const message = vi.fn(async function () {
        messageThis = this;
      });
      dialogApi.confirm = confirm;
      dialogApi.message = message;
      (globalThis as any).__TAURI__ = { plugin: { dialog: dialogApi } };

      const confirmFn = getTauriDialogConfirmOrNull();
      const messageFn = getTauriDialogMessageOrNull();
      expect(confirmFn).not.toBeNull();
      expect(messageFn).not.toBeNull();
      await confirmFn!("Hello");
      await messageFn!("Hello");
      expect(confirmThis).toBe(dialogApi);
      expect(messageThis).toBe(dialogApi);

      const dialogApi2: any = {};
      let confirmThis2: any = null;
      let alertThis2: any = null;
      const confirm2 = vi.fn(async function () {
        confirmThis2 = this;
        return true;
      });
      const alert2 = vi.fn(async function () {
        alertThis2 = this;
      });
      dialogApi2.confirm = confirm2;
      dialogApi2.alert = alert2;
      (globalThis as any).__TAURI__ = { plugins: { dialog: dialogApi2 } };

      const confirmFn2 = getTauriDialogConfirmOrNull();
      const alertFn2 = getTauriDialogMessageOrNull();
      expect(confirmFn2).not.toBeNull();
      expect(alertFn2).not.toBeNull();
      await confirmFn2!("Hello");
      await alertFn2!("Hello");
      expect(confirmThis2).toBe(dialogApi2);
      expect(alertThis2).toBe(dialogApi2);
    });
  });

  describe("getTauriEventApi*", () => {
    it("returns null / throws when the event API is missing", () => {
      expect(getTauriEventApiOrNull()).toBeNull();
      expect(() => getTauriEventApiOrThrow()).toThrowError("Tauri event API not available");
    });

    it("returns listen and a nullable emit (emit missing) and preserves method binding", async () => {
      const eventApi: any = {};
      let listenThis: any = null;
      const handler = vi.fn();
      const listen = vi.fn(async function (_event: string, _handler: (event: any) => void) {
        listenThis = this;
        return () => {};
      });
      eventApi.listen = listen;
      (globalThis as any).__TAURI__ = { event: eventApi };

      const api = getTauriEventApiOrThrow();
      expect(api.emit).toBeNull();
      const unlisten = await api.listen("startup:metrics", handler);
      expect(typeof unlisten).toBe("function");
      expect(listenThis).toBe(eventApi);
      expect(listen).toHaveBeenCalledWith("startup:metrics", handler);
    });

    it("returns listen and emit when both are present and preserves method binding", async () => {
      const eventApi: any = {};
      let listenThis: any = null;
      let emitThis: any = null;
      const handler = vi.fn();
      const listen = vi.fn(async function () {
        listenThis = this;
        return () => {};
      });
      const emit = vi.fn(function () {
        emitThis = this;
      });
      eventApi.listen = listen;
      eventApi.emit = emit;
      (globalThis as any).__TAURI__ = { event: eventApi };

      const api = getTauriEventApiOrThrow();
      await api.listen("startup:metrics", handler);
      await api.emit?.("ping");
      await api.emit?.("ping", { value: 1 });
      expect(listenThis).toBe(eventApi);
      expect(emitThis).toBe(eventApi);
      expect(emit.mock.calls[0]).toEqual(["ping"]);
      expect(emit.mock.calls[1]).toEqual(["ping", { value: 1 }]);
    });

    it("detects the event API under __TAURI__.plugin.event (legacy shape)", async () => {
      const eventApi: any = {};
      let listenThis: any = null;
      let emitThis: any = null;
      const handler = vi.fn();
      const listen = vi.fn(async function () {
        listenThis = this;
        return () => {};
      });
      const emit = vi.fn(function () {
        emitThis = this;
      });
      eventApi.listen = listen;
      eventApi.emit = emit;
      (globalThis as any).__TAURI__ = { plugin: { event: eventApi } };

      const api = getTauriEventApiOrThrow();
      await api.listen("startup:metrics", handler);
      await api.emit?.("ping");
      expect(listenThis).toBe(eventApi);
      expect(emitThis).toBe(eventApi);
    });

    it("detects the event API under __TAURI__.plugins.event (alternate plugin container shape)", async () => {
      const eventApi: any = {};
      let listenThis: any = null;
      let emitThis: any = null;
      const handler = vi.fn();
      const listen = vi.fn(async function () {
        listenThis = this;
        return () => {};
      });
      const emit = vi.fn(function () {
        emitThis = this;
      });
      eventApi.listen = listen;
      eventApi.emit = emit;
      (globalThis as any).__TAURI__ = { plugins: { event: eventApi } };

      const api = getTauriEventApiOrThrow();
      await api.listen("startup:metrics", handler);
      await api.emit?.("ping");
      expect(listenThis).toBe(eventApi);
      expect(emitThis).toBe(eventApi);
    });

    it("treats throwing nested properties (e.g. event getter) as unavailable", () => {
      const tauri: any = {};
      Object.defineProperty(tauri, "event", {
        configurable: true,
        get() {
          throw new Error("Blocked event access");
        },
      });
      (globalThis as any).__TAURI__ = tauri;

      expect(getTauriEventApiOrNull()).toBeNull();
      expect(() => getTauriEventApiOrThrow()).toThrowError("Tauri event API not available");
    });
  });

  describe("getTauriInvoke*", () => {
    it("returns null / throws when the invoke API is missing", () => {
      expect(getTauriInvokeOrNull()).toBeNull();
      expect(hasTauriInvoke()).toBe(false);
      expect(() => getTauriInvokeOrThrow()).toThrowError("Tauri invoke API not available");
    });

    it("detects core.invoke when available and preserves method binding + argument count", async () => {
      const core: any = {};
      let invokeThis: any = null;
      const invoke = vi.fn(async function (cmd: string, args?: any) {
        invokeThis = this;
        return { cmd, args, argc: arguments.length };
      });
      core.invoke = invoke;
      (globalThis as any).__TAURI__ = { core };

      const invokeFn = getTauriInvokeOrNull();
      expect(typeof invokeFn).toBe("function");
      expect(getTauriInvokeOrThrow()).toBe(invokeFn);
      expect(hasTauriInvoke()).toBe(true);

      await expect(invokeFn!("ping")).resolves.toEqual({ cmd: "ping", args: undefined, argc: 1 });
      expect(invokeThis).toBe(core);
      expect(invoke.mock.calls[0]).toEqual(["ping"]);

      await expect(invokeFn!("ping", { value: 1 })).resolves.toEqual({ cmd: "ping", args: { value: 1 }, argc: 2 });
      expect(invoke.mock.calls[1]).toEqual(["ping", { value: 1 }]);

      // Repeated access should be stable for performance (no new wrapper allocation).
      expect(getTauriInvokeOrNull()).toBe(invokeFn);
    });

    it("treats throwing nested properties (e.g. core getter) as unavailable", () => {
      const tauri: any = {};
      Object.defineProperty(tauri, "core", {
        configurable: true,
        get() {
          throw new Error("Blocked core access");
        },
      });
      (globalThis as any).__TAURI__ = tauri;

      expect(getTauriInvokeOrNull()).toBeNull();
      expect(hasTauriInvoke()).toBe(false);
      expect(() => getTauriInvokeOrThrow()).toThrowError("Tauri invoke API not available");
    });
  });

  describe("getTauriAppGetName/Version*", () => {
    it("returns null when the app API is missing", () => {
      expect(getTauriAppGetNameOrNull()).toBeNull();
      expect(getTauriAppGetVersionOrNull()).toBeNull();
    });

    it("detects __TAURI__.app.getName/getVersion when available and preserves method binding", async () => {
      const appApi: any = {};
      const getName = vi.fn(function () {
        return Promise.resolve(this === appApi ? "Formula" : "bad");
      });
      const getVersion = vi.fn(function () {
        return Promise.resolve(this === appApi ? "1.2.3" : "bad");
      });
      appApi.getName = getName;
      appApi.getVersion = getVersion;
      (globalThis as any).__TAURI__ = { app: appApi };

      const nameFn = getTauriAppGetNameOrNull();
      const versionFn = getTauriAppGetVersionOrNull();
      expect(nameFn).not.toBeNull();
      expect(versionFn).not.toBeNull();

      await expect(nameFn!()).resolves.toBe("Formula");
      await expect(versionFn!()).resolves.toBe("1.2.3");
      expect(getName).toHaveBeenCalledTimes(1);
      expect(getVersion).toHaveBeenCalledTimes(1);
    });

    it("throws when the underlying API resolves to a non-string value", async () => {
      const getName = vi.fn(async () => 123);
      const getVersion = vi.fn(async () => null);
      (globalThis as any).__TAURI__ = { app: { getName, getVersion } };

      const nameFn = getTauriAppGetNameOrNull();
      const versionFn = getTauriAppGetVersionOrNull();
      expect(nameFn).not.toBeNull();
      expect(versionFn).not.toBeNull();

      await expect(nameFn!()).rejects.toThrowError("Tauri app.getName returned a non-string value");
      await expect(versionFn!()).rejects.toThrowError("Tauri app.getVersion returned a non-string value");
    });

    it("supports the __TAURI__.plugin.app API shape", async () => {
      const appApi: any = {};
      const getName = vi.fn(async () => "Formula");
      appApi.getName = getName;
      (globalThis as any).__TAURI__ = { plugin: { app: appApi } };

      const nameFn = getTauriAppGetNameOrNull();
      expect(nameFn).not.toBeNull();
      await expect(nameFn!()).resolves.toBe("Formula");
    });

    it("treats throwing nested properties (e.g. app getter) as unavailable", () => {
      const tauri: any = {};
      Object.defineProperty(tauri, "app", {
        configurable: true,
        get() {
          throw new Error("Blocked app access");
        },
      });
      (globalThis as any).__TAURI__ = tauri;

      expect(getTauriAppGetNameOrNull()).toBeNull();
      expect(getTauriAppGetVersionOrNull()).toBeNull();
    });
  });

  describe("getTauriWindowHandle*", () => {
    it("returns null / throws when the window API is missing", () => {
      expect(getTauriWindowHandleOrNull()).toBeNull();
      expect(() => getTauriWindowHandleOrThrow()).toThrowError("Tauri window API not available");
    });

    it("exposes separate feature-detection helpers for the window namespace vs handle resolution", () => {
      expect(hasTauriWindowApi()).toBe(false);
      expect(hasTauriWindowHandleApi()).toBe(false);

      (globalThis as any).__TAURI__ = { window: {} };
      expect(hasTauriWindowApi()).toBe(true);
      expect(hasTauriWindowHandleApi()).toBe(false);

      (globalThis as any).__TAURI__ = { window: { appWindow: { label: "main" } } };
      expect(hasTauriWindowApi()).toBe(true);
      expect(hasTauriWindowHandleApi()).toBe(true);
    });

    it("prefers getCurrentWebviewWindow when available", () => {
      const handle = { label: "main" };
      (globalThis as any).__TAURI__ = {
        window: {
          getCurrentWebviewWindow: vi.fn(() => handle),
          getCurrentWindow: vi.fn(() => ({ label: "fallback" })),
          appWindow: { label: "appWindow" },
        },
      };

      expect(getTauriWindowHandleOrThrow()).toBe(handle);
    });

    it("falls back through getCurrentWindow / getCurrent / appWindow", () => {
      const handle = { label: "appWindow" };
      (globalThis as any).__TAURI__ = {
        window: {
          getCurrentWebviewWindow: vi.fn(() => null),
          getCurrentWindow: vi.fn(() => null),
          getCurrent: vi.fn(() => null),
          appWindow: handle,
        },
      };

      expect(getTauriWindowHandleOrThrow()).toBe(handle);
    });

    it("treats throwing getCurrent* accessors as unavailable and continues probing", () => {
      const handle = { label: "fallback" };
      (globalThis as any).__TAURI__ = {
        window: {
          getCurrentWebviewWindow: vi.fn(() => {
            throw new Error("boom");
          }),
          getCurrentWindow: vi.fn(() => handle),
        },
      };

      expect(getTauriWindowHandleOrNull()).toBe(handle);
      expect(getTauriWindowHandleOrThrow()).toBe(handle);
    });

    it("treats throwing getCurrent* property getters as unavailable and continues probing", () => {
      const handle = { label: "fallback" };
      const winApi: any = { getCurrentWindow: vi.fn(() => handle) };
      Object.defineProperty(winApi, "getCurrentWebviewWindow", {
        configurable: true,
        get() {
          throw new Error("blocked getter");
        },
      });
      (globalThis as any).__TAURI__ = { window: winApi };

      expect(hasTauriWindowHandleApi()).toBe(true);
      expect(getTauriWindowHandleOrNull()).toBe(handle);
      expect(getTauriWindowHandleOrThrow()).toBe(handle);
    });

    it("throws a distinct error when the window API exists but no handle can be resolved", () => {
      (globalThis as any).__TAURI__ = {
        window: {
          getCurrentWebviewWindow: vi.fn(() => null),
          getCurrentWindow: vi.fn(() => null),
          getCurrent: vi.fn(() => null),
          appWindow: null,
        },
      };

      expect(getTauriWindowHandleOrNull()).toBeNull();
      expect(() => getTauriWindowHandleOrThrow()).toThrowError("Tauri window handle not available");
    });

    it("treats throwing nested properties (e.g. window getter) as unavailable", () => {
      const tauri: any = {};
      Object.defineProperty(tauri, "window", {
        configurable: true,
        get() {
          throw new Error("Blocked window access");
        },
      });
      (globalThis as any).__TAURI__ = tauri;

      expect(hasTauriWindowApi()).toBe(false);
      expect(hasTauriWindowHandleApi()).toBe(false);
      expect(getTauriWindowHandleOrNull()).toBeNull();
      expect(() => getTauriWindowHandleOrThrow()).toThrowError("Tauri window API not available");
    });
  });
});
