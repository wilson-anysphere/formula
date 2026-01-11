export const POLICY_SOURCE: Readonly<{
  ORG: string;
  DOCUMENT: string;
  EFFECTIVE: string;
}>;

export function createDefaultOrgPolicy(): any;
export function validatePolicy(policy: any): void;
export function mergePolicies(params: { orgPolicy: any; documentPolicy?: any }): { policy: any; source: any };

