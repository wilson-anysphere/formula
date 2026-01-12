import path from "node:path";
import { loadConfig } from "../config";
import { createPool } from "../db/pool";
import { runMigrations } from "../db/migrations";
import { putSecret, secretExists } from "../secrets/secretStore";
import { totpSecretName } from "../auth/mfa";

async function main(): Promise<void> {
  const config = loadConfig();
  const pool = createPool(config.databaseUrl);
  const migrationsDir = path.resolve(process.cwd(), "migrations");

  try {
    await runMigrations(pool, { migrationsDir });

    const legacy = await pool.query(
      `
        SELECT id, mfa_totp_secret_legacy
        FROM users
        WHERE mfa_totp_secret_legacy IS NOT NULL
      `
    );

    let migrated = 0;

    for (const row of legacy.rows as Array<{ id: string; mfa_totp_secret_legacy: string }>) {
      const userId = String(row.id);
      const secret = row.mfa_totp_secret_legacy;
      if (!secret) continue;

      const name = totpSecretName(userId);

      if (!(await secretExists(pool, name))) {
        await putSecret(pool, config.secretStoreKeys, name, secret);
      }

      await pool.query("UPDATE users SET mfa_totp_secret_legacy = null WHERE id = $1", [userId]);
      migrated += 1;
    }

    // eslint-disable-next-line no-console
    console.log(`migrated ${migrated} TOTP secrets`);
  } finally {
    await pool.end();
  }
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exitCode = 1;
});
