/// <reference types="node" />
/**
 * Tests that `src/` avoids JSDOM dependencies in runtime code.
 *
 * Used by: `npm test` to enforce runtime bundle purity.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Guardrail scan for `jsdom` imports in Worker code paths.
 *
 * Assumptions:
 * - Only `src/` should be free of JSDOM (tests can depend on it).
 */

import assert from "node:assert";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";

const IMPORT_REGEX = /\bfrom\s+["']jsdom["']|\brequire\(["']jsdom["']\)|\bimport\(["']jsdom["']\)/;

/**
 * Recursively collect file paths under a directory.
 *
 * @param dir - Root directory to walk.
 * @param files - Accumulator of file paths.
 * @returns Array of file paths.
 */
const walk = (dir: string, files: string[] = []) => {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      walk(fullPath, files);
    } else {
      files.push(fullPath);
    }
  }
  return files;
};

test("src does not import jsdom (runtime guardrail)", () => {
  const srcRoot = path.join(process.cwd(), "src");
  const files = walk(srcRoot).filter(file => /\.(ts|tsx|js|mjs)$/.test(file));
  const offenders: string[] = [];
  for (const file of files) {
    const content = fs.readFileSync(file, "utf8");
    if (IMPORT_REGEX.test(content)) {
      offenders.push(path.relative(process.cwd(), file));
    }
  }
  assert.strictEqual(offenders.length, 0, `jsdom import found in src: ${offenders.join(", ")}`);
});
