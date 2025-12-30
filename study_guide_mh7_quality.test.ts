/// <reference types="node" />
/**
 * Tests MH7 quality gates for maximal study guide outputs.
 *
 * Used by: `npm test` to validate MH7 validation rules and renderer output.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Placeholder rejection, topic density, discriminator columns, and drug coverage gates.
 *
 * Assumptions:
 * - Uses in-memory HTML strings and JSDOM for structural checks.
 */

import assert from "node:assert";
import test from "node:test";
import { JSDOM } from "jsdom";
import {
  ensureMaximalDiscriminatorColumns,
  ensureMaximalDrugCoverage,
  ensureMaximalPlaceholderQuality,
  ensureMaximalTopicDensity,
} from "./src/study_guides/validate";
import { extractTopicInventoryFromSlides } from "./src/study_guides/inventory";
import type { FactRegistryTopic, FactRegistryFields, FactRegistry } from "./src/study_guides/fact_registry";
import type { TopicInventory } from "./src/study_guides/inventory";
import { renderMaximalStudyGuideHtml } from "./src/study_guides/render_maximal_html";
import { TABLE_SCHEMAS } from "./src/study_guides/table_schemas";

/**
 * Build a FactRegistryFields fixture with optional overrides.
 */
const buildFields = (overrides?: Partial<FactRegistryFields>): FactRegistryFields => ({
  definition_or_role: [],
  mechanism: [],
  clinical_use_indications: [],
  toxicity_adverse_effects: { common: [], serious: [] },
  pk_pearls: [],
  contraindications_warnings: [],
  monitoring: [],
  dosing_regimens_if_given: [],
  interactions_genetics: [],
  ...overrides,
});

test("topic classifier filters headings and keeps real topics", () => {
  const slides = [
    {
      n: 1,
      text: [
        "Objectives",
        "Timeline",
        "'''plaintext",
        "Tacrolimus",
        "Calcineurin inhibitors",
        "Hyperacute rejection",
        "Slide 2",
      ].join("\n"),
    },
  ];
  const inventory = extractTopicInventoryFromSlides(slides);
  assert.ok(inventory.drugs.includes("Tacrolimus"));
  assert.ok(inventory.drug_classes.includes("Calcineurin inhibitors"));
  assert.ok(inventory.conditions.includes("Hyperacute rejection"));
  assert.ok(!inventory.conditions.some(item => /objectives|timeline|plaintext|slide/i.test(item)));
});

test("placeholder gate rejects placeholder phrases", () => {
  assert.throws(
    () => ensureMaximalPlaceholderQuality("<p>Not stated in lecture</p>"),
    /MAXIMAL_QUALITY_FAILED_PLACEHOLDERS/,
  );
  assert.throws(
    () => ensureMaximalPlaceholderQuality("<p>N\\/A</p>"),
    /MAXIMAL_QUALITY_FAILED_PLACEHOLDERS/,
  );
  assert.throws(
    () => ensureMaximalPlaceholderQuality("<p>Not specified</p>"),
    /MAXIMAL_QUALITY_FAILED_PLACEHOLDERS/,
  );
  assert.doesNotThrow(() => ensureMaximalPlaceholderQuality("<p>All good</p>"));
});

test("topic density gate enforces minimum facts", () => {
  const base: FactRegistryTopic = {
    topic_id: "hyperacute_rejection",
    label: "Hyperacute rejection",
    kind: "condition",
    fields: buildFields({
      definition_or_role: [{ text: "antibody-mediated rejection", span_id: "S1" }],
      mechanism: [{ text: "preformed antibodies", span_id: "S2" }],
    }),
  };
  assert.throws(() => ensureMaximalTopicDensity([base]), /MAXIMAL_QUALITY_FAILED_TOPIC_DENSITY/);

  const ok: FactRegistryTopic = {
    ...base,
    fields: buildFields({
      definition_or_role: [{ text: "antibody-mediated rejection", span_id: "S1" }],
      mechanism: [{ text: "preformed antibodies", span_id: "S2" }],
      clinical_use_indications: [{ text: "urgent graft removal", span_id: "S3" }],
    }),
  };
  assert.doesNotThrow(() => ensureMaximalTopicDensity([ok]));
});

test("discriminator column gate enforces schema headers via data-table-id", () => {
  const missingWhy = [
    "<table data-table-id=\"rapid-approach-summary\"><thead><tr>",
    "<th>Clue</th><th>Think of</th><th>Confirm</th><th>Treat</th>",
    "</tr></thead></table>",
    "<table data-table-id=\"treatments-management\"><thead><tr>",
    "<th>Drug/Class</th><th>Use</th><th>Mechanism</th><th>Signature toxicity</th><th>Monitor</th>",
    "</tr></thead></table>",
  ].join("");
  try {
    ensureMaximalDiscriminatorColumns(missingWhy);
    assert.fail("Expected discriminator column failure.");
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    assert.ok(/MAXIMAL_QUALITY_FAILED_DISCRIMINATOR_COLUMNS/.test(message));
    assert.ok(message.includes("detected="));
    assert.ok(message.includes("expected="));
  }

  const ok = [
    "<table data-table-id=\"rapid-approach-summary\"><thead><tr>",
    "<th>Clue</th><th>Think of</th><th>Why</th><th>Confirm</th><th>Treat</th>",
    "</tr></thead></table>",
    "<table data-table-id=\"treatments-management\"><thead><tr>",
    "<th>Drug/Class</th><th>Mechanism</th><th>Toxicity</th><th>Monitoring</th><th>Pearls</th>",
    "</tr></thead></table>",
  ].join("");
  assert.doesNotThrow(() => ensureMaximalDiscriminatorColumns(ok));
});

test("discriminator column gate fails when a required table is missing", () => {
  const html = "<table data-table-id=\"rapid-approach-summary\"><thead><tr><th>Clue</th></tr></thead></table>";
  assert.throws(() => ensureMaximalDiscriminatorColumns(html), /MAXIMAL_QUALITY_FAILED_MISSING_TABLE/);
});

test("drug coverage gate enforces mechanism + tox + PK + use", () => {
  const base: FactRegistryTopic = {
    topic_id: "tacrolimus",
    label: "Tacrolimus",
    kind: "drug",
    fields: buildFields(),
  };
  assert.throws(() => ensureMaximalDrugCoverage([base]), /MAXIMAL_QUALITY_FAILED_DRUG_COVERAGE/);

  const ok: FactRegistryTopic = {
    ...base,
    fields: buildFields({
      mechanism: [{ text: "inhibits calcineurin", span_id: "S1" }],
      toxicity_adverse_effects: {
        common: [
          { text: "nephrotoxicity", span_id: "S2" },
          { text: "neurotoxicity", span_id: "S3" },
        ],
        serious: [],
      },
      pk_pearls: [{ text: "CYP3A metabolism", span_id: "S4" }],
      clinical_use_indications: [{ text: "maintenance immunosuppression", span_id: "S5" }],
    }),
  };
  assert.doesNotThrow(() => ensureMaximalDrugCoverage([ok]));
});

test("maximal renderer emits schema table ids and headers", () => {
  const topic: FactRegistryTopic = {
    topic_id: "tacrolimus",
    label: "Tacrolimus",
    kind: "drug",
    fields: buildFields({
      mechanism: [{ text: "inhibits calcineurin", span_id: "S1" }],
      clinical_use_indications: [{ text: "maintenance immunosuppression", span_id: "S2" }],
      toxicity_adverse_effects: {
        common: [{ text: "nephrotoxicity", span_id: "S3" }],
        serious: [{ text: "neurotoxicity", span_id: "S4" }],
      },
      pk_pearls: [{ text: "CYP3A metabolism", span_id: "S5" }],
      monitoring: [{ text: "monitor trough levels", span_id: "S6" }],
    }),
  };
  const registry: FactRegistry = { topics: [topic], spans: [] };
  const inventory: TopicInventory = {
    conditions: ["Tacrolimus"],
    drugs: ["Tacrolimus"],
    drug_classes: [],
    phenotypes: [],
    processes: [],
    garbage: [],
    tests: [],
    treatments: [],
    formulas_cutoffs: [],
    mechanisms: [],
  };
  const html = renderMaximalStudyGuideHtml({
    lectureTitle: "Test Lecture",
    buildUtc: "1970-01-01T00:00:00Z",
    slideCount: 1,
    slides: [{ n: 1, text: "Tacrolimus" }],
    inventory,
    registry,
  });
  const dom = new JSDOM(html);
  const rapidTable = dom.window.document.querySelector('table[data-table-id="rapid-approach-summary"]');
  const treatmentTable = dom.window.document.querySelector('table[data-table-id="treatments-management"]');
  assert.ok(rapidTable);
  assert.ok(treatmentTable);
  const rapidHeaders = Array.from(rapidTable!.querySelectorAll("th")).map(th => (th.textContent || "").trim());
  const treatmentHeaders = Array.from(treatmentTable!.querySelectorAll("th")).map(th => (th.textContent || "").trim());
  assert.deepStrictEqual(rapidHeaders, TABLE_SCHEMAS["rapid-approach-summary"].requiredHeaders);
  assert.deepStrictEqual(treatmentHeaders, TABLE_SCHEMAS["treatments-management"].requiredHeaders);
  const rapidRows = Array.from(rapidTable!.querySelectorAll("tbody tr"));
  const treatmentRows = Array.from(treatmentTable!.querySelectorAll("tbody tr"));
  assert.ok(rapidRows.length >= 1);
  assert.ok(treatmentRows.length >= 1);
  const appendixIdx = html.indexOf("slide-by-slide-appendix");
  const rapidIdx = html.indexOf("data-table-id=\"rapid-approach-summary\"");
  const treatmentIdx = html.indexOf("data-table-id=\"treatments-management\"");
  assert.ok(rapidIdx !== -1 && treatmentIdx !== -1);
  if (appendixIdx !== -1) {
    assert.ok(rapidIdx < appendixIdx);
    assert.ok(treatmentIdx < appendixIdx);
  }
});
