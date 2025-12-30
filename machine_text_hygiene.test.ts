/// <reference types="node" />
/**
 * Tests text hygiene helpers for slide/document cleaning.
 *
 * Used by: `npm test` to validate sanitization logic.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Refusal-line stripping and placeholder insertion for empty slides.
 *
 * Assumptions:
 * - Slide markers follow the "Slide N (p.N):" convention.
 */

import assert from "node:assert";
import test from "node:test";
import { sanitizeDocText } from "./src/machine/text_hygiene";

test("sanitizeDocText removes refusal lines and preserves slides", () => {
  const input = [
    "Slide 1 (p.1):",
    "I can't assist with that.",
    "",
    "Slide 2 (p.2):",
    "Valid content line",
  ].join("\n");

  const output = sanitizeDocText(input);

  assert.strictEqual(/can't assist/i.test(output), false);
  assert.strictEqual(output.includes("[NO TEXT]"), true);
  assert.strictEqual(output.includes("Valid content line"), true);
});
