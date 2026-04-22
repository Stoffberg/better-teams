import { parentPort, workerData } from "node:worker_threads";
import {
  extractTokens,
  getAuthToken,
  getAvailableAccounts,
  getCachedPresence,
} from "./token-store";

type TokenWorkerRequest =
  | { operation: "extractTokens" }
  | { operation: "getAuthToken"; tenantId?: string | null }
  | { operation: "getAvailableAccounts" }
  | { operation: "getCachedPresence"; userMris: string[] };

function runRequest(request: TokenWorkerRequest): unknown {
  switch (request.operation) {
    case "extractTokens":
      return extractTokens();
    case "getAuthToken":
      return getAuthToken(request.tenantId);
    case "getAvailableAccounts":
      return getAvailableAccounts();
    case "getCachedPresence":
      return getCachedPresence(request.userMris);
  }
}

try {
  parentPort?.postMessage({ ok: true, value: runRequest(workerData) });
} catch (error) {
  parentPort?.postMessage({
    ok: false,
    error: error instanceof Error ? error.message : String(error),
  });
}
