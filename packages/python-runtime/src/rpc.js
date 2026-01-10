/**
 * Generic spreadsheet API RPC dispatcher.
 *
 * Both the native Python runtime (stdio) and the Pyodide runtime (worker)
 * ultimately need the same behavior: given an `api` object and a method name,
 * invoke the appropriate host function.
 *
 * The API can be provided as:
 * - An object with methods (e.g. `api.get_cell_value({ ... })`)
 * - An object with a generic `call(method, params)` function
 */
export async function dispatchRpc(api, method, params) {
  if (api && typeof api[method] === "function") {
    return await api[method](params);
  }
  if (api && typeof api.call === "function") {
    return await api.call(method, params);
  }
  throw new Error(`Spreadsheet API does not implement RPC method "${method}"`);
}

