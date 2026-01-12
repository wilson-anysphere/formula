export class YjsVersionStore {
  [key: string]: any;
  constructor(...args: any[]);

  saveVersion(version: any): Promise<void>;
  getVersion(versionId: string): Promise<any | null>;
  listVersions(): Promise<any[]>;
  updateVersion(versionId: string, patch: any): Promise<void>;
  deleteVersion(versionId: string): Promise<void>;
}
