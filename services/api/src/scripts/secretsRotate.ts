import { loadConfig } from "../config";
import { createPool } from "../db/pool";
import { runSecretsRotation } from "../secrets/rotation";

const config = loadConfig();
const pool = createPool(config.databaseUrl);

const batchSize = process.env.BATCH_SIZE ? Number.parseInt(process.env.BATCH_SIZE, 10) : undefined;
const prefix = process.env.PREFIX?.trim() || undefined;

runSecretsRotation(pool, config.secretStoreKeys, { batchSize, prefix })
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
