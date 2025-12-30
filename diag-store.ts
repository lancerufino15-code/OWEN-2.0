/**
 * Purpose:
 * - Diagnostic KV helpers for capturing last library ask and analytics keys.
 *
 * Responsibilities:
 * - Resolve an available KV binding, persist lightweight debug snapshots, and read them back for admin endpoints.
 *
 * Architecture role:
 * - Shared utility used by the worker to surface recent operations for troubleshooting.
 *
 * Constraints:
 * - Cloudflare Workers KV only; logs once when no KV binding is available.
 */
import type { Env } from "./types";

const KEY_LAST_ASK = "diag/LAST_LIBRARY_ASK.json";
const KEY_LAST_ANALYTICS = "diag/LAST_ANALYTICS_KEY.txt";
let loggedMissingDiagKv = false;

function getDiagKv(env: Env): KVNamespace | null {
  return env.OWEN_DIAG_KV || env.DOCS_KV || null;
}

function logMissingKvOnce() {
  if (loggedMissingDiagKv) return;
  loggedMissingDiagKv = true;
  console.warn("[DIAG_KV_MISSING] no KV binding available");
}

/**
 * Return the KV binding name available for diagnostics.
 *
 * @param env - Worker environment with KV bindings.
 * @returns Binding name or null when none are configured.
 */
export function diagKvBindingName(env: Env): string | null {
  if (env.OWEN_DIAG_KV) return "OWEN_DIAG_KV";
  if (env.DOCS_KV) return "DOCS_KV";
  return null;
}

/**
 * Persist the last `/api/library/ask` payload snapshot for diagnostics.
 *
 * @param env - Worker environment with KV bindings.
 * @param snapshot - Arbitrary snapshot object to store.
 * @remarks Side effects: writes to KV; logs when KV is missing or write fails.
 */
export async function saveLastLibraryAsk(env: Env, snapshot: any) {
  const kv = getDiagKv(env);
  if (!kv) {
    logMissingKvOnce();
    return;
  }
  try {
    await kv.put(KEY_LAST_ASK, JSON.stringify(snapshot), { expirationTtl: 60 * 60 * 24 * 3 }); // keep a few days
  } catch (err) {
    console.warn("[DIAG_KV_WRITE_FAILED]", { key: KEY_LAST_ASK, err: err instanceof Error ? err.message : String(err) });
  }
}

/**
 * Load the most recent `/api/library/ask` snapshot from diagnostics KV.
 *
 * @param env - Worker environment with KV bindings.
 * @returns Parsed snapshot or null on any failure.
 */
export async function loadLastLibraryAsk(env: Env): Promise<any | null> {
  const kv = getDiagKv(env);
  if (!kv) return null;
  try {
    const text = await kv.get(KEY_LAST_ASK);
    if (!text) return null;
    return JSON.parse(text);
  } catch {
    return null;
  }
}

/**
 * Persist the last analytics object key for debugging admin endpoints.
 *
 * @param env - Worker environment with KV bindings.
 * @param key - Analytics object key to store.
 * @remarks Side effects: writes to KV; logs when KV is missing or write fails.
 */
export async function saveLastAnalyticsKey(env: Env, key: string | null) {
  const kv = getDiagKv(env);
  if (!kv) {
    logMissingKvOnce();
    return;
  }
  if (!key) return;
  try {
    await kv.put(KEY_LAST_ANALYTICS, key, { expirationTtl: 60 * 60 * 24 * 7 });
  } catch (err) {
    console.warn("[DIAG_KV_WRITE_FAILED]", { key: KEY_LAST_ANALYTICS, err: err instanceof Error ? err.message : String(err) });
  }
}

/**
 * Load the last stored analytics object key from diagnostics KV.
 *
 * @param env - Worker environment with KV bindings.
 * @returns Analytics key or null on any failure.
 */
export async function loadLastAnalyticsKey(env: Env): Promise<string | null> {
  const kv = getDiagKv(env);
  if (!kv) return null;
  try {
    const text = await kv.get(KEY_LAST_ANALYTICS);
    return text || null;
  } catch {
    return null;
  }
}
