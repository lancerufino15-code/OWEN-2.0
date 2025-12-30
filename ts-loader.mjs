/**
 * Node ESM loader that resolves bare `./path` to `.ts`/`.js` during tests.
 *
 * Used by: `npm test` via `--loader ./ts-loader.mjs` to allow TypeScript imports.
 *
 * Key exports:
 * - `resolve`: Node ESM loader hook for extensionless specifiers.
 *
 * Assumptions:
 * - Only relative specifiers without extensions are rewritten; others delegate to Node.
 */
import { stat } from "node:fs/promises";
import { extname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

async function fileExists(url) {
  try {
    await stat(fileURLToPath(url));
    return true;
  } catch {
    return false;
  }
}

/**
 * Resolve module specifiers without extensions to `.ts` or `.js` files.
 *
 * @param specifier - Import specifier string.
 * @param context - Loader context with parent URL.
 * @param nextResolve - Fallback resolver.
 * @returns Resolver result for Node's loader hook.
 */
export async function resolve(specifier, context, nextResolve) {
  if ((specifier.startsWith("./") || specifier.startsWith("../")) && !extname(specifier)) {
    const parentURL = context.parentURL || pathToFileURL(`${process.cwd()}/`).href;
    const tsUrl = new URL(`${specifier}.ts`, parentURL);
    if (await fileExists(tsUrl)) {
      return { url: tsUrl.href, shortCircuit: true };
    }
    const jsUrl = new URL(`${specifier}.js`, parentURL);
    if (await fileExists(jsUrl)) {
      return { url: jsUrl.href, shortCircuit: true };
    }
    const indexTsUrl = new URL(`${specifier}/index.ts`, parentURL);
    if (await fileExists(indexTsUrl)) {
      return { url: indexTsUrl.href, shortCircuit: true };
    }
    const indexJsUrl = new URL(`${specifier}/index.js`, parentURL);
    if (await fileExists(indexJsUrl)) {
      return { url: indexJsUrl.href, shortCircuit: true };
    }
  }
  return nextResolve(specifier, context);
}
