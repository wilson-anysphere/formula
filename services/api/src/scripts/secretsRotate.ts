import { loadConfig } from "../config";
import { createPool } from "../db/pool";
import { runSecretsRotation } from "../secrets/rotation";

const config = loadConfig();
const pool = createPool(config.databaseUrl);

const batchSizeRaw = process.env.BATCH_SIZE?.trim();
const batchSizeParsed = batchSizeRaw ? Number.parseInt(batchSizeRaw, 10) : undefined;
const batchSize =
  batchSizeParsed != null && Number.isFinite(batchSizeParsed) && batchSizeParsed > 0
    ? batchSizeParsed
    : undefined;
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
