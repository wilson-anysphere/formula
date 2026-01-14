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
  hasTauri,
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

    it("detects the dialog API on __TAURI__.dialog (legacy shape)", () => {
      const open = vi.fn();
      const save = vi.fn();
      (globalThis as any).__TAURI__ = { dialog: { open, save } };

      const api = getTauriDialogOrThrow();
      expect(api.open).toBe(open);
      expect(api.save).toBe(save);
    });

    it("detects the dialog API on __TAURI__.plugin.dialog (tauri v2 plugin shape)", () => {
      const open = vi.fn();
      const save = vi.fn();
      (globalThis as any).__TAURI__ = { plugin: { dialog: { open, save } } };

      const api = getTauriDialogOrThrow();
      expect(api.open).toBe(open);
      expect(api.save).toBe(save);
    });

    it("detects the dialog API on __TAURI__.plugins.dialog (alternate plugin container shape)", () => {
      const open = vi.fn();
      const save = vi.fn();
      (globalThis as any).__TAURI__ = { plugins: { dialog: { open, save } } };

      const api = getTauriDialogOrThrow();
      expect(api.open).toBe(open);
      expect(api.save).toBe(save);
    });

    it("treats partial dialog APIs as unavailable (e.g. open without save)", () => {
      const open = vi.fn();
      (globalThis as any).__TAURI__ = { dialog: { open } };
      expect(getTauriDialogOrNull()).toBeNull();
      expect(() => getTauriDialogOrThrow()).toThrowError("Tauri dialog API not available");
    });

    it("exposes open/save-only helpers that do not require the full API surface", () => {
      const open = vi.fn();
      (globalThis as any).__TAURI__ = { dialog: { open } };

      expect(getTauriDialogOpenOrNull()).toBe(open);
      expect(getTauriDialogSaveOrNull()).toBeNull();
    });

    it("detects confirm() when available", () => {
      const confirm = vi.fn();
      (globalThis as any).__TAURI__ = { dialog: { confirm } };
      expect(getTauriDialogConfirmOrNull()).toBe(confirm);
    });

    it("detects message() (or alert()) when available", () => {
      const message = vi.fn();
      (globalThis as any).__TAURI__ = { dialog: { message } };
      expect(getTauriDialogMessageOrNull()).toBe(message);

      const alert = vi.fn();
      (globalThis as any).__TAURI__ = { dialog: { alert } };
      expect(getTauriDialogMessageOrNull()).toBe(alert);
    });
  });

  describe("getTauriEventApi*", () => {
    it("returns null / throws when the event API is missing", () => {
      expect(getTauriEventApiOrNull()).toBeNull();
      expect(() => getTauriEventApiOrThrow()).toThrowError("Tauri event API not available");
    });

    it("returns listen and a nullable emit (emit missing)", () => {
      const listen = vi.fn(async () => () => {});
      (globalThis as any).__TAURI__ = { event: { listen } };

      const api = getTauriEventApiOrThrow();
      expect(api.listen).toBe(listen);
      expect(api.emit).toBeNull();
    });

    it("returns listen and emit when both are present", () => {
      const listen = vi.fn(async () => () => {});
      const emit = vi.fn();
      (globalThis as any).__TAURI__ = { event: { listen, emit } };

      const api = getTauriEventApiOrThrow();
      expect(api.listen).toBe(listen);
      expect(api.emit).toBe(emit);
    });

    it("detects the event API under __TAURI__.plugin.event (legacy shape)", () => {
      const listen = vi.fn(async () => () => {});
      const emit = vi.fn();
      (globalThis as any).__TAURI__ = { plugin: { event: { listen, emit } } };

      const api = getTauriEventApiOrThrow();
      expect(api.listen).toBe(listen);
      expect(api.emit).toBe(emit);
    });

    it("detects the event API under __TAURI__.plugins.event (alternate plugin container shape)", () => {
      const listen = vi.fn(async () => () => {});
      const emit = vi.fn();
      (globalThis as any).__TAURI__ = { plugins: { event: { listen, emit } } };

      const api = getTauriEventApiOrThrow();
      expect(api.listen).toBe(listen);
      expect(api.emit).toBe(emit);
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
  });
});
