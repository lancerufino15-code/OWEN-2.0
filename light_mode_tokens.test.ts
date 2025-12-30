/// <reference types="node" />
/**
 * Tests light-mode token defaults used by the UI.
 *
 * Used by: `npm test` to enforce theme design constraints.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Ensures light-mode token blocks avoid pure white (#fff/#ffffff).
 *
 * Assumptions:
 * - CSS tokens are declared in a `:root { ... }` block in HTML/TS sources.
 */

import assert from "node:assert";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

const PURE_WHITE_RE = /#(?:fff|ffffff)\b/i;

/**
 * Extract the :root CSS block from a source file.
 *
 * @param source - Raw file contents.
 * @param label - Label used for assertion messages.
 * @returns The :root block contents.
 */
function extractRootBlock(source: string, label: string): string {
  const match = source.match(/:root\s*{([^}]*)}/s);
  assert.ok(match, `${label} is missing a :root token block.`);
  return match[1];
}

/**
 * Assert that a CSS block contains no pure white tokens.
 *
 * @param block - CSS block text.
 * @param label - Label used for assertion messages.
 */
function assertNoPureWhite(block: string, label: string): void {
  assert.ok(!PURE_WHITE_RE.test(block), `${label} contains pure white (#fff/#ffffff) in light-mode tokens.`);
}

test("light mode tokens avoid pure white in public UI", () => {
  const html = readFileSync(path.join(process.cwd(), "public/index.html"), "utf8");
  const rootBlock = extractRootBlock(html, "public/index.html");
  assertNoPureWhite(rootBlock, "public/index.html");
});

test("light mode tokens avoid pure white in study guide CSS", () => {
  const contracts = readFileSync(path.join(process.cwd(), "src/study_guides/contracts.ts"), "utf8");
  const rootBlock = extractRootBlock(contracts, "src/study_guides/contracts.ts");
  assertNoPureWhite(rootBlock, "src/study_guides/contracts.ts");
});
