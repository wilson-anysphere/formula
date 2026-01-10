import path from "node:path";
import { loadConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { createPool } from "../db/pool";

const config = loadConfig();
const pool = createPool(config.databaseUrl);

const migrationsDir = path.resolve(process.cwd(), "migrations");

runMigrations(pool, { migrationsDir })
  .then(async () => {
    await pool.end();
    // eslint-disable-next-line no-console
    console.log("migrations applied");
  })
  .catch(async (err) => {
    // eslint-disable-next-line no-console
    console.error(err);
    await pool.end();
    process.exitCode = 1;
  });
