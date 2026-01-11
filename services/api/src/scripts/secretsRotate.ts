import { loadConfig } from "../config";
import { createPool } from "../db/pool";
import { runSecretsRotation } from "../secrets/rotation";

const config = loadConfig();
const pool = createPool(config.databaseUrl);

runSecretsRotation(pool, config.secretStoreKeys)
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

