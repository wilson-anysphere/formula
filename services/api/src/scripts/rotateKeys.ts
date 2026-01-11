import { loadConfig } from "../config";
import { createKeyring } from "../crypto/keyring";
import { runKeyRotation } from "../crypto/rotation";
import { createPool } from "../db/pool";

const config = loadConfig();
const pool = createPool(config.databaseUrl);
const keyring = createKeyring(config);

runKeyRotation(pool, keyring)
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

