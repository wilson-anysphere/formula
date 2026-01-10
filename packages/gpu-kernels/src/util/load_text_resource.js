/**
 * Loads a text resource (WGSL) in both browser and Node.
 * @param {URL} url
 */
export async function loadTextResource(url) {
  if (url.protocol === "file:") {
    const fs = await import("node:fs/promises");
    return fs.readFile(url, "utf8");
  }

  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to fetch resource: ${url} (${res.status})`);
  }
  return res.text();
}

