/**
 * Tests markdown table parsing helpers used in the chat UI.
 *
 * Used by: `npm test` to validate table parsing rules.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Divider handling, header inference, escaped pipes, and trailing text capture.
 *
 * Assumptions:
 * - Input blocks are individual markdown table candidates.
 */
import assert from "node:assert/strict";
import { test } from "node:test";
import { parseMarkdownTableBlock } from "./public/table_parser.mjs";

test("parses markdown table with divider", () => {
  const block = [
    "| Name | Value |",
    "| --- | --- |",
    "| Alpha | One |",
    "| Beta | Two |",
  ].join("\n");
  const parsed = parseMarkdownTableBlock(block);
  assert.ok(parsed);
  assert.equal(parsed.hasHeader, true);
  assert.deepEqual(parsed.headers, ["Name", "Value"]);
  assert.equal(parsed.rows.length, 2);
  assert.equal(parsed.columnCount, 2);
});

test("parses headerless pipe table with inferred step headers", () => {
  const block = [
    "| Step 1 | Do the first thing |",
    "| Step 2 | Do the next thing |",
    "| Step 3 | Finish the flow |",
  ].join("\n");
  const parsed = parseMarkdownTableBlock(block);
  assert.ok(parsed);
  assert.equal(parsed.columnCount, 2);
  assert.equal(parsed.rows.length, 3);
  assert.equal(parsed.hasHeader, true);
  assert.deepEqual(parsed.headers, ["Step", "Details"]);
});

test("parses headerless pipe table without inferred headers", () => {
  const block = [
    "| Alpha | First value |",
    "| Beta | Second value |",
  ].join("\n");
  const parsed = parseMarkdownTableBlock(block);
  assert.ok(parsed);
  assert.equal(parsed.hasHeader, false);
  assert.deepEqual(parsed.headers, []);
});

test("ignores fenced code blocks with pipes", () => {
  const block = [
    "```bash",
    "| a | b |",
    "| c | d |",
    "```",
  ].join("\n");
  assert.equal(parseMarkdownTableBlock(block), null);
});

test("rejects unstable pipe tables", () => {
  const block = [
    "| A | B |",
    "| C |",
  ].join("\n");
  assert.equal(parseMarkdownTableBlock(block), null);
});

test("parses markdown table with heading caption", () => {
  const block = [
    "## Antibiotic dosing table",
    "| Drug | Dose |",
    "| --- | --- |",
    "| Amoxicillin | 500 mg |",
    "| Doxycycline | 100 mg |",
  ].join("\n");
  const parsed = parseMarkdownTableBlock(block);
  assert.ok(parsed);
  assert.equal(parsed.caption, "## Antibiotic dosing table");
  assert.equal(parsed.columnCount, 2);
});

test("preserves trailing text after a divider table", () => {
  const block = [
    "| A | B |",
    "| --- | --- |",
    "| 1 | 2 |",
    "Plan:",
    "- Do X",
  ].join("\n");
  const parsed = parseMarkdownTableBlock(block);
  assert.ok(parsed);
  assert.equal(parsed.rows.length, 1);
  assert.equal(parsed.trailingText, "Plan:\n- Do X");
});

test("keeps escaped pipes inside cells", () => {
  const block = [
    "| Name | Value |",
    "| --- | --- |",
    "| Foo \\| Bar | Baz |",
    "| Qux | Quux |",
  ].join("\n");
  const parsed = parseMarkdownTableBlock(block);
  assert.ok(parsed);
  assert.equal(parsed.columnCount, 2);
  assert.equal(parsed.rows[0][0], "Foo | Bar");
});
