export const MEDIA: {
  prefersDark: string;
  forcedColors: string;
  prefersContrastMore: string;
  reducedMotion: string;
};

export function getSystemTheme(env?: any): "light" | "dark" | "high-contrast";
export function getSystemReducedMotion(env?: any): boolean;
export function subscribeToMediaQuery(env: any, query: string, onChange: (matches: boolean) => void): () => void;

