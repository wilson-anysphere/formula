export class InMemoryAuditLogger {
  events: any[];
  log(event: any): string;
  list(): any[];
}

