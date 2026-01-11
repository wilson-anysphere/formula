import { loadConfig } from "../config";
import { createKeyring } from "../crypto/keyring";
import { encryptPlaintextDocumentVersions } from "../db/documentVersions";
import { createPool } from "../db/pool";

const config = loadConfig();
const pool = createPool(config.databaseUrl);
const keyring = createKeyring(config);

const orgId = process.env.ORG_ID;
const batchSize = process.env.BATCH_SIZE ? Number.parseInt(process.env.BATCH_SIZE, 10) : 100;

encryptPlaintextDocumentVersions(pool, keyring, { orgId, batchSize })
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

