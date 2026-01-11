import { loadConfig } from "../config";
import { runKmsRotationSweep } from "../crypto/kms";
import { createPool } from "../db/pool";

const config = loadConfig();
const pool = createPool(config.databaseUrl);

runKmsRotationSweep(pool)
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
