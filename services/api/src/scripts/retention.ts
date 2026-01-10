import { loadConfig } from "../config";
import { createPool } from "../db/pool";
import { runRetentionSweep } from "../retention";

const config = loadConfig();
const pool = createPool(config.databaseUrl);

runRetentionSweep(pool)
  .then(async (result) => {
    await pool.end();
    // eslint-disable-next-line no-console
    console.log(JSON.stringify(result));
  })
  .catch(async (err) => {
    // eslint-disable-next-line no-console
    console.error(err);
    await pool.end();
    process.exitCode = 1;
  });

