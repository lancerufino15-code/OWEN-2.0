
/**
 * Module declaration shims for `.mjs` imports in the TS build.
 *
 * Used by: root worker entrypoint and tests when TypeScript resolves `.mjs` imports.
 *
 * Key exports:
 * - Ambient module declarations for the PDF.js legacy wrapper and generic `.mjs` imports.
 *
 * Assumptions:
 * - Types here are intentionally loose stubs; runtime behavior comes from the actual `.mjs` bundles.
 */
declare module "./pdfjs-dist-legacy-build-pdf.mjs" {
  export const __isStub: boolean;
  export const GlobalWorkerOptions: { workerSrc?: string };
  export function getDocument(...args: any[]): any;
  const _default: any;
  export default _default;
}

declare module "*.mjs" {
  const value: any;
  export default value;
}
