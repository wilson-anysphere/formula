const { InMemoryObjectStore } = require("./inMemoryObjectStore.js");

class RegionalObjectStore {
  constructor({ regions } = {}) {
    this._stores = new Map();
    if (regions) {
      for (const region of regions) {
        this._stores.set(region, new InMemoryObjectStore());
      }
    }
  }

  _ensure(region) {
    if (!this._stores.has(region)) {
      this._stores.set(region, new InMemoryObjectStore());
    }
    return this._stores.get(region);
  }

  async putObject(region, key, value) {
    await this._ensure(region).putObject(key, value);
  }

  async getObject(region, key) {
    return this._ensure(region).getObject(key);
  }

  async deleteObject(region, key) {
    await this._ensure(region).deleteObject(key);
  }

  storeForRegion(region) {
    return this._ensure(region);
  }
}

module.exports = { RegionalObjectStore };
