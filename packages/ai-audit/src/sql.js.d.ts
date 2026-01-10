declare module "sql.js" {
  export interface SqlJsStatic {
    Database: new (data?: Uint8Array) => any;
  }

  export interface InitSqlJsConfig {
    locateFile?: (file: string) => string;
  }

  export default function initSqlJs(config?: InitSqlJsConfig): Promise<SqlJsStatic>;
}

