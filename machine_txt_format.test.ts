/// <reference types="node" />
/**
 * Tests machine TXT formatting for extracted PDF text.
 *
 * Used by: `npm test` to validate PDF page normalization into slide text.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Page ordering, line trimming, and missing-page placeholders.
 *
 * Assumptions:
 * - Input text uses the `--- Page N ---` markers produced by PDF normalization.
 */

import assert from "node:assert";
import test from "node:test";
import { formatMachineTxtFromExtractedText } from "./src/index";

test("formatMachineTxtFromExtractedText orders pages, trims lines, and fills gaps", () => {
  const input = [
    "--- Page 2 ---",
    "Line B  ",
    "Line B-2   ",
    "",
    "--- Page 1 ---",
    "Line A",
    "",
    "--- Page 4 ---",
    "",
  ].join("\n");

  const result = formatMachineTxtFromExtractedText(input, 4);

  const expected = [
    "Slide 1 (p.1):",
    "Line A",
    "",
    "Slide 2 (p.2):",
    "Line B",
    "Line B-2",
    "",
    "Slide 3 (p.3): [NO TEXT]",
    "",
    "Slide 4 (p.4): [NO TEXT]",
  ].join("\n");

  assert.strictEqual(result, expected);
  assert.ok(!result.includes("\r"));
});
