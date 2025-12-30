/**
 * Public re-exports for study guide validation helpers.
 *
 * Used by: `src/index.ts` and tests to import runtime validators from a stable path.
 *
 * Key exports:
 * - Re-exported validators from `validate.runtime.ts` (structure, coverage, style).
 *
 * Assumptions:
 * - This barrel stays in sync with `validate.runtime.ts` to provide a stable import path.
 */
export * from "./validate.runtime.ts";
