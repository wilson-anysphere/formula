export class InMemoryKeychainProvider {
  constructor() {
    this._secrets = new Map();
  }

  _key(service, account) {
    return `${service}:${account}`;
  }

  async getSecret({ service, account }) {
    const value = this._secrets.get(this._key(service, account));
    return value ? Buffer.from(value) : null;
  }

  async setSecret({ service, account, secret }) {
    if (!Buffer.isBuffer(secret)) {
      throw new TypeError("secret must be a Buffer");
    }
    this._secrets.set(this._key(service, account), Buffer.from(secret));
  }

  async deleteSecret({ service, account }) {
    this._secrets.delete(this._key(service, account));
  }
}

