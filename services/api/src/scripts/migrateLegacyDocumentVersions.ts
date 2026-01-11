import { loadConfig } from "../config";
import { KmsProviderFactory } from "../crypto/kms";
import { migrateLegacyEncryptedDocumentVersions } from "../db/documentVersions";
import { createPool } from "../db/pool";

const config = loadConfig();
const pool = createPool(config.databaseUrl);
const kmsFactory = new KmsProviderFactory(pool, {
  aws: { enabled: config.awsKmsEnabled, region: config.awsRegion ?? null }
});

const orgId = process.env.ORG_ID;
const batchSize = process.env.BATCH_SIZE ? Number.parseInt(process.env.BATCH_SIZE, 10) : 100;

migrateLegacyEncryptedDocumentVersions(pool, kmsFactory, {
  orgId,
  batchSize,
  legacyLocalKmsMasterKey: config.localKmsMasterKey
})
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
