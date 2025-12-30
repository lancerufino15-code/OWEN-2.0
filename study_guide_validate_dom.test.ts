/// <reference types="node" />
/**
 * Tests DOM validation helpers for study guide HTML output.
 *
 * Used by: `npm test` to validate DOM-based schema checks.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Required table presence and discriminator header validation.
 *
 * Assumptions:
 * - Uses JSDOM to parse HTML fragments without external assets.
 */

import assert from "node:assert";
import test from "node:test";
import { JSDOM } from "jsdom";
import { TABLE_SCHEMA_LIST, type TableSchema } from "./src/study_guides/table_schemas";

/**
 * Normalize table header text for comparison.
 */
const normalizeHeaderLabel = (label: string) =>
  (label || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

/**
 * Check whether required headers match a table's normalized header list.
 */
const matchesRequiredHeader = (schema: TableSchema, required: string, normalizedHeaders: string[]) => {
  const aliases = schema.allowedHeaderAliases?.[required] || [];
  const candidates = [required, ...aliases].map(normalizeHeaderLabel).filter(Boolean);
  return candidates.some(candidate => normalizedHeaders.includes(candidate));
};

/**
 * Validate discriminator columns in maximal tables using DOM parsing.
 */
const ensureMaximalDiscriminatorColumnsDom = (html: string) => {
  const dom = new JSDOM(html || "");
  const doc = dom.window.document;
  const missingTables: string[] = [];
  const headerFailures: Array<{ id: string; missing: string[] }> = [];

  for (const schema of TABLE_SCHEMA_LIST) {
    if (!schema.mustExistInMaximal) continue;
    const table = doc.querySelector(`table[data-table-id="${schema.id}"]`);
    if (!table) {
      missingTables.push(schema.id);
      continue;
    }
    const detected = Array.from(table.querySelectorAll("th"))
      .map(th => (th.textContent || "").trim())
      .filter(Boolean)
      .map(normalizeHeaderLabel);
    const missing = schema.requiredHeaders.filter(required => !matchesRequiredHeader(schema, required, detected));
    if (missing.length) {
      headerFailures.push({ id: schema.id, missing });
    }
  }

  if (missingTables.length) {
    throw new Error(`MAXIMAL_QUALITY_FAILED_MISSING_TABLE: ${missingTables.join(", ")}`);
  }
  if (headerFailures.length) {
    const details = headerFailures.map(failure => `${failure.id} missing=${failure.missing.join("|")}`).join(" ; ");
    throw new Error(`MAXIMAL_QUALITY_FAILED_DISCRIMINATOR_COLUMNS: ${details}`);
  }
};

test("dom validator catches missing discriminator columns", () => {
  const html = [
    "<table data-table-id=\"rapid-approach-summary\"><thead><tr>",
    "<th>Clue</th><th>Think of</th><th>Confirm</th><th>Treat</th>",
    "</tr></thead></table>",
    "<table data-table-id=\"treatments-management\"><thead><tr>",
    "<th>Drug/Class</th><th>Use</th><th>Mechanism</th><th>Signature toxicity</th><th>Monitor</th>",
    "</tr></thead></table>",
  ].join("");
  assert.throws(
    () => ensureMaximalDiscriminatorColumnsDom(html),
    /MAXIMAL_QUALITY_FAILED_DISCRIMINATOR_COLUMNS/,
  );
});
