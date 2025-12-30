/**
 * Module declaration shims for `.mjs` imports in the TS build.
 *
 * Used by: the Worker TypeScript build to allow importing MJS assets.
 *
 * Key exports:
 * - Ambient module declarations for the PDF.js legacy wrapper and generic `.mjs` modules.
 *
 * Assumptions:
 * - These are loose type shims only; real runtime behavior comes from the bundled MJS files.
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
