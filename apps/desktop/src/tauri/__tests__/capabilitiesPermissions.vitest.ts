import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

describe("Tauri capabilities", () => {
  it("allowlists add_sheet so the sheet tabs '+' button can persist sheets in Tauri builds", () => {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };

    expect(Array.isArray(capability.permissions)).toBe(true);

    const permissions = capability.permissions as unknown[];
    const allowInvoke = permissions.find(
      (permission) =>
        Boolean(permission) && typeof permission === "object" && (permission as Record<string, unknown>).identifier === "core:allow-invoke",
    );

    expect(allowInvoke).toBeTruthy();
    const allowedCommands = (allowInvoke as Record<string, unknown>).allow;
    expect(Array.isArray(allowedCommands)).toBe(true);

    const commands = allowedCommands as unknown[];
    expect(commands).toContain("add_sheet");
  });

  it("explicitly allowlists rich clipboard commands required by the desktop clipboard provider", () => {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };

    expect(Array.isArray(capability.permissions)).toBe(true);

    const permissions = capability.permissions as unknown[];
    const allowInvoke = permissions.find(
      (permission) => Boolean(permission) && typeof permission === "object" && (permission as Record<string, unknown>).identifier === "core:allow-invoke",
    );

    expect(allowInvoke).toBeTruthy();
    const allowedCommands = (allowInvoke as Record<string, unknown>).allow;
    expect(Array.isArray(allowedCommands)).toBe(true);

    const commands = allowedCommands as unknown[];
    expect(commands).toContain("clipboard_read");
    expect(commands).toContain("clipboard_write");
    // Back-compat for older desktop builds.
    expect(commands).toContain("read_clipboard");
    expect(commands).toContain("write_clipboard");
  });

  it("allows invoking show_system_notification from the main window", () => {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };

    expect(Array.isArray(capability.permissions)).toBe(true);

    const permissions = capability.permissions as unknown[];
    const allowInvoke = permissions.find(
      (permission) => Boolean(permission) && typeof permission === "object" && (permission as Record<string, unknown>).identifier === "core:allow-invoke",
    );

    expect(allowInvoke).toBeTruthy();
    const allowedCommands = (allowInvoke as Record<string, unknown>).allow;
    expect(Array.isArray(allowedCommands)).toBe(true);

    const commands = allowedCommands as unknown[];
    expect(commands).toContain("show_system_notification");
  });

  it("allows invoking oauth_loopback_listen for RFC 8252 loopback redirect capture", () => {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };

    expect(Array.isArray(capability.permissions)).toBe(true);

    const permissions = capability.permissions as unknown[];
    const allowInvoke = permissions.find(
      (permission) => Boolean(permission) && typeof permission === "object" && (permission as Record<string, unknown>).identifier === "core:allow-invoke",
    );

    expect(allowInvoke).toBeTruthy();
    const allowedCommands = (allowInvoke as Record<string, unknown>).allow;
    expect(Array.isArray(allowedCommands)).toBe(true);

    const commands = allowedCommands as unknown[];
    expect(commands).toContain("oauth_loopback_listen");
  });

  it("routes external URL opening through open_external_url (no shell:allow-open permission)", () => {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };

    expect(Array.isArray(capability.permissions)).toBe(true);

    const permissions = capability.permissions as unknown[];
    const allowInvoke = permissions.find(
      (permission) => Boolean(permission) && typeof permission === "object" && (permission as Record<string, unknown>).identifier === "core:allow-invoke",
    );

    expect(allowInvoke).toBeTruthy();
    const allowedCommands = (allowInvoke as Record<string, unknown>).allow;
    expect(Array.isArray(allowedCommands)).toBe(true);

    const commands = allowedCommands as unknown[];
    expect(commands).toContain("open_external_url");
    expect(permissions).not.toContain("shell:allow-open");
  });

  it("grants the updater permissions required by the frontend restart/install flow", () => {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };

    expect(Array.isArray(capability.permissions)).toBe(true);

    const permissions = capability.permissions as unknown[];

    const allowInvoke = permissions.find(
      (permission) =>
        Boolean(permission) && typeof permission === "object" && (permission as Record<string, unknown>).identifier === "core:allow-invoke",
    );
    expect(allowInvoke).toBeTruthy();
    const allowedCommands = (allowInvoke as Record<string, unknown>).allow;
    expect(Array.isArray(allowedCommands)).toBe(true);

    const commands = allowedCommands as unknown[];
    // The updater restart flow calls `restart_app` when available, falling back to `quit_app`.
    expect(commands).toContain("restart_app");
    expect(commands).toContain("quit_app");

    expect(permissions).toContain("updater:allow-check");
    expect(permissions).toContain("updater:allow-download");
    expect(permissions).toContain("updater:allow-install");
    // Ensure we keep the updater permission surface minimal/explicit.
    expect(permissions).not.toContain("updater:default");
    expect(permissions).not.toContain("updater:allow-download-and-install");
  });
});
