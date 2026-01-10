import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

function ensureString(value, name) {
  if (typeof value !== "string" || value.length === 0) {
    throw new TypeError(`${name} must be a non-empty string`);
  }
}

export class OsKeychainProvider {
  constructor({ platform = process.platform } = {}) {
    this.platform = platform;
  }

  async getSecret({ service, account }) {
    ensureString(service, "service");
    ensureString(account, "account");

    if (this.platform === "darwin") {
      try {
        const { stdout } = await execFileAsync("security", [
          "find-generic-password",
          "-s",
          service,
          "-a",
          account,
          "-w"
        ]);
        return Buffer.from(stdout.trim(), "utf8");
      } catch (error) {
        // Not found is a normal condition.
        return null;
      }
    }

    // Secret Service via libsecret / secret-tool.
    if (this.platform === "linux") {
      try {
        const { stdout } = await execFileAsync("secret-tool", [
          "lookup",
          "service",
          service,
          "account",
          account
        ]);
        return Buffer.from(stdout.trim(), "utf8");
      } catch (error) {
        return null;
      }
    }

    throw new Error(`OS keychain not implemented for platform: ${this.platform}`);
  }

  async setSecret({ service, account, secret }) {
    ensureString(service, "service");
    ensureString(account, "account");
    if (!Buffer.isBuffer(secret)) {
      throw new TypeError("secret must be a Buffer");
    }

    if (this.platform === "darwin") {
      await execFileAsync("security", [
        "add-generic-password",
        "-U",
        "-s",
        service,
        "-a",
        account,
        "-w",
        secret.toString("utf8")
      ]);
      return;
    }

    if (this.platform === "linux") {
      // secret-tool does not support non-interactive store without stdin piping.
      // We avoid plumbing stdin here; production desktop should use a native
      // keychain library instead of CLI wrappers.
      throw new Error(
        "Linux Secret Service setSecret not implemented (use native bindings in production)"
      );
    }

    throw new Error(`OS keychain not implemented for platform: ${this.platform}`);
  }

  async deleteSecret({ service, account }) {
    ensureString(service, "service");
    ensureString(account, "account");

    if (this.platform === "darwin") {
      try {
        await execFileAsync("security", [
          "delete-generic-password",
          "-s",
          service,
          "-a",
          account
        ]);
      } catch (error) {
        // Ignore if missing.
      }
      return;
    }

    if (this.platform === "linux") {
      throw new Error(
        "Linux Secret Service deleteSecret not implemented (use native bindings in production)"
      );
    }

    throw new Error(`OS keychain not implemented for platform: ${this.platform}`);
  }
}

