import assert from "node:assert/strict";
import test from "node:test";

import { SharePointConnector } from "../../src/connectors/sharepoint.js";

/**
 * @param {any} data
 * @param {number} [status]
 * @param {Record<string, string>} [headers]
 */
function jsonResponse(data, status = 200, headers = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json", ...headers },
  });
}

test("SharePointConnector: SharePoint.Contents lists drives and injects OAuth2 bearer (retry on 401)", async () => {
  /** @type {any[]} */
  const tokenCalls = [];

  const oauth2Manager = {
    getAccessToken: async (opts) => {
      tokenCalls.push(opts);
      return { accessToken: opts.forceRefresh ? "token-2" : "token-1", expiresAtMs: null, refreshToken: null };
    },
  };

  const siteEndpoint = "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Finance?$select=id,lastModifiedDateTime,webUrl";
  const drivesEndpoint = "https://graph.microsoft.com/v1.0/sites/site-id/drives?$select=id,name,webUrl,driveType";

  /** @type {Array<{ url: string; auth?: string }>} */
  const fetchCalls = [];
  /** @type {typeof fetch} */
  const fetchFn = async (url, init) => {
    const auth = /** @type {any} */ (init?.headers)?.Authorization;
    fetchCalls.push({ url: String(url), auth });

    if (url === siteEndpoint) {
      if (auth === "Bearer token-1") {
        return new Response("unauthorized", { status: 401 });
      }
      assert.equal(auth, "Bearer token-2");
      return jsonResponse(
        { id: "site-id", webUrl: "https://contoso.sharepoint.com/sites/Finance", lastModifiedDateTime: "2024-01-01T00:00:00Z" },
        200,
        { etag: "W/\"123\"" },
      );
    }

    if (url === drivesEndpoint) {
      assert.equal(auth, "Bearer token-2");
      return jsonResponse({
        value: [{ id: "drive-1", name: "Documents", webUrl: "https://contoso.sharepoint.com/sites/Finance/Shared Documents", driveType: "documentLibrary" }],
      });
    }

    throw new Error(`Unexpected URL: ${url}`);
  };

  const connector = new SharePointConnector({ fetch: fetchFn, oauth2Manager });
  const result = await connector.execute({
    siteUrl: "https://Contoso.SharePoint.com/sites/Finance/",
    mode: "contents",
    options: { auth: { type: "oauth2", providerId: "example", scopes: ["Sites.Read.All"] } },
  });

  assert.deepEqual(result.table.toGrid(), [
    ["Name", "Id", "WebUrl", "DriveType"],
    ["Documents", "drive-1", "https://contoso.sharepoint.com/sites/Finance/Shared Documents", "documentLibrary"],
  ]);

  assert.equal(tokenCalls.length, 2);
  assert.equal(tokenCalls[0].forceRefresh, false);
  assert.equal(tokenCalls[1].forceRefresh, true);
  assert.equal(fetchCalls.filter((c) => c.url === siteEndpoint).length, 2, "expected site resolution retry");
});

test("SharePointConnector: SharePoint.Files paginates and can recurse folders", async () => {
  const siteEndpoint = "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Finance?$select=id,lastModifiedDateTime,webUrl";
  const drivesEndpoint = "https://graph.microsoft.com/v1.0/sites/site-id/drives?$select=id,name,webUrl,driveType";
  const rootChildrenEndpoint =
    "https://graph.microsoft.com/v1.0/drives/drive-1/root/children?$select=id,name,webUrl,size,file,folder,parentReference,lastModifiedDateTime,createdDateTime";
  const rootChildrenNext = "https://graph.microsoft.com/v1.0/drives/drive-1/root/children?page=2";
  const folderChildrenEndpoint =
    "https://graph.microsoft.com/v1.0/drives/drive-1/items/folder-1/children?$select=id,name,webUrl,size,file,folder,parentReference,lastModifiedDateTime,createdDateTime";

  /** @type {Set<string>} */
  const seenUrls = new Set();
  /** @type {typeof fetch} */
  const fetchFn = async (url, _init) => {
    seenUrls.add(String(url));

    if (url === siteEndpoint) {
      return jsonResponse({ id: "site-id", webUrl: "https://contoso.sharepoint.com/sites/Finance", lastModifiedDateTime: "2024-01-01T00:00:00Z" });
    }

    if (url === drivesEndpoint) {
      return jsonResponse({
        value: [{ id: "drive-1", name: "Documents", webUrl: "https://contoso.sharepoint.com/sites/Finance/Shared Documents", driveType: "documentLibrary" }],
      });
    }

    if (url === rootChildrenEndpoint) {
      return jsonResponse({
        value: [
          { id: "folder-1", name: "Folder", folder: { childCount: 1 } },
          { id: "file-1", name: "a.txt", file: { mimeType: "text/plain" }, size: 10, webUrl: "https://file/a.txt" },
        ],
        "@odata.nextLink": rootChildrenNext,
      });
    }

    if (url === rootChildrenNext) {
      return jsonResponse({
        value: [{ id: "file-2", name: "b.csv", file: { mimeType: "text/csv" }, size: 20, webUrl: "https://file/b.csv" }],
      });
    }

    if (url === folderChildrenEndpoint) {
      return jsonResponse({
        value: [{ id: "file-3", name: "c.xlsx", file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }, size: 30 }],
      });
    }

    throw new Error(`Unexpected URL: ${url}`);
  };

  const connector = new SharePointConnector({ fetch: fetchFn });
  const result = await connector.execute({
    siteUrl: "https://contoso.sharepoint.com/sites/Finance",
    mode: "files",
    options: { recursive: true, includeContent: false },
  });

  const nameIndex = result.table.getColumnIndex("Name");
  const names = result.table.rows.map((r) => r[nameIndex]).filter((v) => typeof v === "string");
  names.sort();
  assert.deepEqual(names, ["a.txt", "b.csv", "c.xlsx"]);

  assert.ok(seenUrls.has(rootChildrenNext), "expected pagination request via @odata.nextLink");
  assert.ok(seenUrls.has(folderChildrenEndpoint), "expected recursive folder traversal");
});

