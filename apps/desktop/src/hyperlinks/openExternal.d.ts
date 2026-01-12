export type OpenExternalHyperlinkDeps = {
  shellOpen: (uri: string) => Promise<void>;
  confirmUntrustedProtocol?: (message: string) => Promise<boolean>;
  permissions?: { request: (permission: string, context: any) => Promise<boolean> };
  allowedProtocols?: Set<string>;
};

/**
 * Open an external hyperlink via the host OS.
 * @returns Whether the link was opened.
 */
export function openExternalHyperlink(uri: string, deps: OpenExternalHyperlinkDeps): Promise<boolean>;

