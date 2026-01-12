const fs = require("node:fs/promises");
const path = require("node:path");

/**
 * Best-effort atomic file write (write to a temporary file then rename).
 *
 * This helps avoid partially-written JSON stores if the process crashes or is
 * interrupted mid-write.
 *
 * @param {string} finalPath
 * @param {string | Uint8Array} data
 * @param {BufferEncoding | undefined} [encoding]
 */
async function writeFileAtomic(finalPath, data, encoding) {
  const tmpPath = path.join(
    path.dirname(finalPath),
    `${path.basename(finalPath)}.tmp-${Date.now()}-${Math.random().toString(16).slice(2)}`
  );

  try {
    await fs.writeFile(tmpPath, data, encoding);
    try {
      await fs.rename(tmpPath, finalPath);
    } catch (err) {
      // On Windows, rename does not reliably overwrite existing files.
      if (err && typeof err === "object" && "code" in err && (err.code === "EEXIST" || err.code === "EPERM")) {
        await fs.rm(finalPath, { force: true });
        await fs.rename(tmpPath, finalPath);
        return;
      }
      throw err;
    }
  } catch (err) {
    await fs.rm(tmpPath, { force: true }).catch(() => {});
    throw err;
  }
}

module.exports = {
  writeFileAtomic
};

