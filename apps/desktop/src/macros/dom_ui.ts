import { DefaultMacroSecurityController } from "./security";
import type { MacroBackend, MacroCellUpdate, MacroSecurityStatus } from "./types";
import { MacroRunner } from "./runner";

export interface MacroRunnerRenderOptions {
  onApplyUpdates?: (updates: MacroCellUpdate[]) => void | Promise<void>;
  /**
   * Optional spreadsheet edit-state predicate.
   *
   * When omitted, `renderMacroRunner` falls back to the desktop shell's global
   * `__formulaSpreadsheetIsEditing` flag (when present) and otherwise assumes editing is off.
   */
  isEditing?: (() => boolean) | null;
  /**
   * Optional read-only predicate (e.g. collab viewer/commenter roles).
   *
   * When omitted, read-only is assumed to be false.
   */
  isReadOnly?: (() => boolean) | null;
}

/**
 * Minimal "macro runner" UI using DOM APIs (framework-agnostic).
 *
 * Production builds can wrap this in React or another UI layer; this file exists
 * to encode the core UX requirements in executable code:
 * - Select a macro
 * - Run it
 * - Show output and errors
 * - Prompt for macro enablement (and permissions) when needed
 */
export async function renderMacroRunner(
  container: HTMLElement,
  backend: MacroBackend,
  workbookId: string,
  opts: MacroRunnerRenderOptions = {}
): Promise<void> {
  // Abort any prior run's event listeners so rerendering does not leak window handlers.
  const prevAbort = (container as any).__formulaMacroRunnerAbort as AbortController | undefined;
  try {
    prevAbort?.abort();
  } catch {
    // ignore
  }
  const abort = new AbortController();
  (container as any).__formulaMacroRunnerAbort = abort;

  const security = new DefaultMacroSecurityController();
  const runner = new MacroRunner(backend, security);
  const macros = await runner.list(workbookId);
  const securityStatus = await backend.getMacroSecurityStatus(workbookId);

  container.innerHTML = "";
  container.classList.add("macros-runner");

  const header = document.createElement("div");
  header.className = "macros-runner__header";
  header.textContent = "Macros";

  const securityBanner = document.createElement("div");
  securityBanner.dataset["testid"] = "macro-security-banner";
  securityBanner.className = "macros-runner__security-banner";

  const controls = document.createElement("div");
  controls.className = "macros-runner__controls";

  const select = document.createElement("select");
  select.dataset["testid"] = "macro-runner-select";
  select.className = "macros-runner__select";
  for (const macro of macros) {
    const opt = document.createElement("option");
    opt.value = macro.id;
    opt.textContent = macro.name;
    select.appendChild(opt);
  }

  const trustButton = document.createElement("button");
  trustButton.dataset["testid"] = "macro-runner-trust-center";
  trustButton.type = "button";
  trustButton.className = "macros-runner__button";
  trustButton.textContent = "Trust Center…";

  const runButton = document.createElement("button");
  runButton.dataset["testid"] = "macro-runner-run";
  runButton.type = "button";
  runButton.className = "macros-runner__button";
  runButton.textContent = "Run";

  const output = document.createElement("pre");
  output.className = "macros-runner__output";

  let currentSecurity = securityStatus;
  let running = false;
  let isEditing = (() => {
    if (typeof opts.isEditing === "function") {
      try {
        return Boolean(opts.isEditing());
      } catch {
        return false;
      }
    }
    const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
    return globalEditing === true;
  })();
  let isReadOnly = (() => {
    if (typeof opts.isReadOnly === "function") {
      try {
        return Boolean(opts.isReadOnly());
      } catch {
        return false;
      }
    }
    return false;
  })();
  let mutationsDisabled = isEditing || isReadOnly;

  const syncRunButtonDisabledState = () => {
    runButton.disabled = running || mutationsDisabled;
    if (mutationsDisabled) {
      runButton.title = isReadOnly
        ? "Read-only: you don't have permission to run macros."
        : isEditing
          ? "Finish editing to run a macro."
          : "";
    } else {
      runButton.title = "";
    }
  };

  function renderSecurityBanner(status: MacroSecurityStatus): void {
    securityBanner.dataset["blocked"] = "false";
    if (!status.hasMacros) {
      securityBanner.textContent = "Security: No VBA macros detected.";
      trustButton.disabled = true;
      return;
    }

    const signature = status.signature?.status ?? "unknown";
    const signer = status.signature?.signerSubject ? ` (${status.signature.signerSubject})` : "";
    const origin = status.originPath ? `\nWorkbook: ${status.originPath}` : "";

    securityBanner.textContent = `Security: Trust Center = ${status.trust}\nSignature: ${signature}${signer}${origin}`;

    const signatureOk = signature === "signed_verified" || signature === "signed_untrusted";
    const blocked = status.trust === "blocked" || (status.trust === "trusted_signed_only" && !signatureOk);
    trustButton.disabled = false;
    securityBanner.dataset["blocked"] = blocked ? "true" : "false";
    if (blocked) {
      securityBanner.textContent += "\n\nMacros blocked by Trust Center. Click “Trust Center…” to change this.";
    }
  }

  renderSecurityBanner(currentSecurity);
  syncRunButtonDisabledState();

  const addWindowListener = (type: string, listener: (event: Event) => void) => {
    if (typeof window === "undefined" || typeof window.addEventListener !== "function") return;
    window.addEventListener(type, listener as EventListener);
    abort.signal.addEventListener(
      "abort",
      () => {
        try {
          window.removeEventListener(type, listener as EventListener);
        } catch {
          // ignore
        }
      },
      { once: true },
    );
  };

  addWindowListener("formula:spreadsheet-editing-changed", (evt: Event) => {
    const detail = (evt as CustomEvent)?.detail as any;
    if (detail && typeof detail.isEditing === "boolean") {
      isEditing = detail.isEditing;
    } else {
      const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
      isEditing = globalEditing === true;
    }
    mutationsDisabled = isEditing || isReadOnly;
    syncRunButtonDisabledState();
  });

  addWindowListener("formula:read-only-changed", (evt: Event) => {
    const detail = (evt as CustomEvent)?.detail as any;
    if (detail && typeof detail.readOnly === "boolean") {
      isReadOnly = detail.readOnly;
    } else if (typeof opts.isReadOnly === "function") {
      try {
        isReadOnly = Boolean(opts.isReadOnly());
      } catch {
        isReadOnly = false;
      }
    } else {
      isReadOnly = false;
    }
    mutationsDisabled = isEditing || isReadOnly;
    syncRunButtonDisabledState();
  });

  trustButton.onclick = () => {
    void (async () => {
      try {
        const decision = await security.requestTrustDecision({
          workbookId,
          macroId: select.value,
          status: currentSecurity,
        });
        if (!decision) return;
        currentSecurity = await backend.setMacroTrust(workbookId, decision);
        renderSecurityBanner(currentSecurity);
      } catch (err) {
        output.textContent += `Error: ${String(err)}\n`;
      }
    })().catch((err) => {
      // Best-effort: avoid unhandled rejections from fire-and-forget DOM handlers.
      output.textContent += `Error: ${String(err)}\n`;
    });
  };

  runButton.onclick = () => {
    void (async () => {
      if (mutationsDisabled) return;
      output.textContent = "";
      running = true;
      try {
        syncRunButtonDisabledState();
      } catch {
        // ignore
      }
      try {
        const macroId = select.value;
        const selected = macros.find((m) => m.id === macroId);
        // Script macros can pay a cold-start cost (worker startup, TS transpilation, Pyodide init),
        // so use a more forgiving timeout than the macro runner defaults.
        const timeoutMs = selected?.language === "python" ? 60_000 : 20_000;
        const result = await runner.run({ workbookId, macroId, timeoutMs });
        if (result.output.length) {
          output.textContent += result.output.join("\n") + "\n";
        }
        if (result.updates && result.updates.length) {
          if (opts.onApplyUpdates) {
            try {
              await opts.onApplyUpdates(result.updates);
              output.textContent += `Applied ${result.updates.length} updates.\n`;
            } catch (err) {
              output.textContent += `Error applying updates: ${String(err)}\n`;
            }
          } else {
            output.textContent += `Macro returned ${result.updates.length} updates (not applied).\n`;
          }
        }
        if (result.error?.blocked) {
          output.textContent += `Blocked by Trust Center (${result.error.blocked.reason}).\n`;
        }
        if (!result.ok) {
          output.textContent += `Error: ${result.error?.message ?? "Unknown error"}\n`;
        }
        currentSecurity = await backend.getMacroSecurityStatus(workbookId);
        renderSecurityBanner(currentSecurity);
      } catch (err) {
        output.textContent += `Error: ${String(err)}\n`;
      } finally {
        running = false;
        try {
          syncRunButtonDisabledState();
        } catch {
          // ignore
        }
      }
    })().catch((err) => {
      // Best-effort: avoid unhandled rejections from fire-and-forget DOM handlers.
      running = false;
      try {
        syncRunButtonDisabledState();
      } catch {
        // ignore
      }
      output.textContent += `Error: ${String(err)}\n`;
    });
  };

  controls.appendChild(select);
  controls.appendChild(trustButton);
  controls.appendChild(runButton);

  container.appendChild(header);
  container.appendChild(securityBanner);
  container.appendChild(controls);
  container.appendChild(output);
}
