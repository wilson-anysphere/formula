import { useEffect, useState } from "react";

import {
  clearReservedRootGuardError,
  reservedRootGuardUiMessage,
  subscribeToReservedRootGuardDisconnect,
} from "./collabReservedRootGuard.js";

export { clearReservedRootGuardError };

export function useReservedRootGuardError(provider: any | null): string | null {
  const [detected, setDetected] = useState(false);

  useEffect(() => {
    if (!provider) {
      setDetected(false);
      return;
    }
    return subscribeToReservedRootGuardDisconnect(provider, (nextDetected) => {
      setDetected(nextDetected);
    });
  }, [provider]);

  return detected ? reservedRootGuardUiMessage() : null;
}

