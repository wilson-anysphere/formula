import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";

import {
  getTauriDialogOrNull,
  getTauriDialogOrThrow,
  getTauriEventApiOrNull,
  getTauriEventApiOrThrow,
  hasTauriWindowApi,
  hasTauriWindowHandleApi,
  getTauriWindowHandleOrNull,
  getTauriWindowHandleOrThrow,
} from "../api";

describe("tauri/api dynamic accessors", () => {
  const originalTauri = (globalThis as any).__TAURI__;

  beforeEach(() => {
    (globalThis as any).__TAURI__ = undefined;
  });

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    vi.restoreAllMocks();
  });

  describe("getTauriDialog*", () => {
    it("returns null / throws when __TAURI__ is missing", () => {
      expect(getTauriDialogOrNull()).toBeNull();
      expect(() => getTauriDialogOrThrow()).toThrowError("Tauri dialog API not available");
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
