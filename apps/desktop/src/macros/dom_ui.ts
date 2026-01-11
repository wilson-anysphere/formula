import { DefaultMacroSecurityController } from "./security";
import type { MacroBackend, MacroCellUpdate, MacroSecurityStatus } from "./types";
import { MacroRunner } from "./runner";

export interface MacroRunnerRenderOptions {
  onApplyUpdates?: (updates: MacroCellUpdate[]) => void | Promise<void>;
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
  const security = new DefaultMacroSecurityController();
  const runner = new MacroRunner(backend, security);
  const macros = await runner.list(workbookId);
  const securityStatus = await backend.getMacroSecurityStatus(workbookId);

  container.innerHTML = "";

  const header = document.createElement("div");
  header.textContent = "Macros";

  const securityBanner = document.createElement("div");
  securityBanner.dataset["testid"] = "macro-security-banner";
  securityBanner.style.whiteSpace = "pre-wrap";
  securityBanner.style.marginBottom = "8px";

  const select = document.createElement("select");
  for (const macro of macros) {
    const opt = document.createElement("option");
    opt.value = macro.id;
    opt.textContent = macro.name;
    select.appendChild(opt);
  }

  const trustButton = document.createElement("button");
  trustButton.textContent = "Trust Center…";
  trustButton.style.marginLeft = "8px";

  const runButton = document.createElement("button");
  runButton.textContent = "Run";

  const output = document.createElement("pre");
  output.style.whiteSpace = "pre-wrap";

  let currentSecurity = securityStatus;

  function renderSecurityBanner(status: MacroSecurityStatus): void {
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
    if (blocked) {
      securityBanner.textContent += "\n\nMacros blocked by Trust Center. Click “Trust Center…” to change this.";
    }
  }

  renderSecurityBanner(currentSecurity);

  trustButton.onclick = async () => {
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
  };

  runButton.onclick = async () => {
    output.textContent = "";
    runButton.disabled = true;
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
      runButton.disabled = false;
    }
  };

  container.appendChild(header);
  container.appendChild(securityBanner);
  container.appendChild(select);
  container.appendChild(trustButton);
  container.appendChild(runButton);
  container.appendChild(output);
}
