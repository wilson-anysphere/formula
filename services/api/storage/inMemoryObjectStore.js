class InMemoryObjectStore {
  constructor() {
    this._objects = new Map();
  }

  async putObject(key, value) {
    if (typeof key !== "string" || key.length === 0) {
      throw new TypeError("key must be a non-empty string");
    }
    if (!Buffer.isBuffer(value)) {
      throw new TypeError("value must be a Buffer");
    }
    this._objects.set(key, Buffer.from(value));
  }

  async getObject(key) {
    const value = this._objects.get(key);
    return value ? Buffer.from(value) : null;
  }

  async deleteObject(key) {
    this._objects.delete(key);
  }

  async keys() {
    return Array.from(this._objects.keys());
  }
}

module.exports = { InMemoryObjectStore };
