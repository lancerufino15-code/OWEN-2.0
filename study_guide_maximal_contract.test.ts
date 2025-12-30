/// <reference types="node" />
/**
 * Tests canonical/maximal study guide HTML contract validation.
 *
 * Used by: `npm test` to validate maximal study guide structure and CSS contract.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Required sections, highlight/table classes, and coverage gate behavior.
 *
 * Assumptions:
 * - HTML fixtures are constructed in-memory; no external rendering required.
 */

import assert from "node:assert";
import test from "node:test";
import { JSDOM } from "jsdom";
import { renderStudyGuideHtml, type StepAOutput, type StepBOutput, type StepCOutput } from "./src/machine/render_study_guide_html";
import { BASE_STUDY_GUIDE_CSS, CANONICAL_STYLE_CONTRACT, renderLegend } from "./src/study_guides/contracts";
import type { TopicInventory } from "./src/study_guides/inventory";
import { stripLeadingNumbering } from "./src/study_guides/normalize";
import { ensureMaximalCoverage, validateMaximalStructure, validateStyleContract } from "./src/study_guides/validate";

/**
 * Build a topic inventory fixture for maximal contract tests.
 */
const buildInventory = (): TopicInventory => ({
  conditions: [
    "4. Acute Kidney Injury in Children",
    "6. Pediatric UTIs",
    "8. Electrolyte Disorders in Pediatrics",
    "Diabetic Ketoacidosis",
    "COPD Exacerbation",
    "Salicylate Overdose",
    "Ethylene Glycol Toxicity",
    "Methanol Toxicity",
    "Distal RTA",
    "Proximal RTA",
    "Hyperkalemic RTA",
    "Bartter Syndrome",
    "Gitelman Syndrome",
    "Liddle Syndrome",
    "Hyperaldosteronism",
  ],
  drugs: [],
  drug_classes: [],
  phenotypes: [],
  processes: [],
  garbage: [],
  tests: ["ABG", "Anion Gap"],
  treatments: ["IV fluids", "Insulin"],
  formulas_cutoffs: ["Winter's Formula", "Anion Gap"],
  mechanisms: ["Bicarbonate Buffer"],
});

/**
 * Build a Core Conditions bullet list entry for a condition.
 */
const buildCoreBullet = (condition: string) => {
  const label = stripLeadingNumbering(condition);
  return [
    "<li>",
    `  <strong><span class="hl disease">${label}</span></strong>`,
    "  <ul>",
    `    <li>Context: <span class="hl disease">${label} role</span></li>`,
    "    <li>Key clue: <span class=\"hl symptom\">Example clue</span></li>",
    "    <li>Confirm/Monitor: <span class=\"hl diagnostic\">Example confirm</span></li>",
    "    <li>Treat/Next: <span class=\"hl treatment\">Example treatment</span></li>",
    "  </ul>",
    "</li>",
  ].join("");
};

/**
 * Build a Condition Coverage table row for a condition.
 */
const buildCoverageRow = (condition: string) => {
  const label = stripLeadingNumbering(condition);
  return [
    "<tr>",
    `  <td><span class="hl disease">${label}</span></td>`,
    "  <td><span class=\"hl symptom\">Key clue</span></td>",
    "  <td><span class=\"hl mechanism\">Why discriminator</span></td>",
    "  <td><span class=\"hl diagnostic\">Confirm monitor</span></td>",
    "  <td><span class=\"hl treatment\">Treat next</span></td>",
    "</tr>",
  ].join("");
};

/**
 * Build a maximal HTML fixture using the provided inventory.
 */
const buildMaximalFixtureHtml = (inventory: TopicInventory) => {
  const coreList = inventory.conditions.map(buildCoreBullet).join("");
  const coverageRows = inventory.conditions.map(buildCoverageRow).join("");
  const appendixItems = inventory.conditions.map(condition => `<li>${stripLeadingNumbering(condition)}</li>`).join("");
  return [
    "<!doctype html>",
    "<html>",
    "<head>",
    "<meta charset=\"utf-8\" />",
    "<style>",
    BASE_STUDY_GUIDE_CSS,
    "</style>",
    "</head>",
    "<body>",
    "<header class=\"sticky-header\">Study Guide - Maximal Build</header>",
    "<main class=\"content\">",
    "<section id=\"highlight-legend\">",
    renderLegend(),
    "</section>",
    "<section id=\"core-conditions\">",
    "<h1>Core Conditions &amp; Patterns</h1>",
    `<ul>${coreList}</ul>`,
    "</section>",
    "<section id=\"condition-coverage\">",
    "<h1>Condition Coverage Table</h1>",
    "<table class=\"compare\">",
    "<thead><tr><th>Condition</th><th>Key clue</th><th>Why (discriminator)</th><th>Confirm/Monitor</th><th>Treat/Next step</th></tr></thead>",
    `<tbody>${coverageRows}</tbody>`,
    "</table>",
    "</section>",
    "<section id=\"rapid-approach-summary\">",
    "<table class=\"tri\"><thead><tr><th>Clue</th><th>Think of</th><th>Why (discriminator)</th><th>Confirm/Monitor</th><th>Treat/Next step</th></tr></thead><tbody><tr><td><span class=\"hl symptom\">Clue</span></td><td><span class=\"hl disease\">Dx</span></td><td><span class=\"hl mechanism\">Why</span></td><td><span class=\"hl diagnostic\">Confirm</span></td><td><span class=\"hl treatment\">Treat</span></td></tr></tbody></table>",
    "</section>",
    "<section id=\"differential-diagnosis\">",
    "<table class=\"compare\"><thead><tr><th>Type</th><th>Timing</th><th>Why (discriminator)</th><th>Key implication</th></tr></thead><tbody><tr><td><span class=\"hl disease\">Acute</span></td><td><span class=\"hl symptom\">Days</span></td><td><span class=\"hl mechanism\">Cellular</span></td><td><span class=\"hl treatment\">Treat</span></td></tr></tbody></table>",
    "</section>",
    "<section id=\"cutoffs-formulas\">",
    "<table class=\"cutoff\"><thead><tr><th>Item</th><th>Value</th><th>Note</th></tr></thead><tbody><tr><td><span class=\"hl cutoff\">Anion gap</span></td><td><span class=\"hl cutoff\">&gt; 12</span></td><td>High</td></tr></tbody></table>",
    "</section>",
    "<section id=\"slide-by-slide-appendix\">",
    `<ul>${appendixItems}</ul>`,
    "</section>",
    "<section id=\"coverage-qa\">",
    "<h1>Coverage &amp; QA</h1>",
    "<table class=\"cutoff\"><thead><tr><th>Inventory</th><th>Count</th></tr></thead><tbody><tr><td>Conditions</td><td>1</td></tr></tbody></table>",
    "<p>Missing items: none</p>",
    "</section>",
    "</main>",
    "</body>",
    "</html>",
  ].join("\n");
};

test("style contract exists in maximal fixture output", () => {
  const inventory = buildInventory();
  const html = buildMaximalFixtureHtml(inventory);
  const styleCheck = validateStyleContract(html);
  assert.strictEqual(styleCheck.ok, true);
  const structureCheck = validateMaximalStructure(html);
  assert.strictEqual(structureCheck.ok, true);

  const dom = new JSDOM(html);
  const legend = dom.window.document.querySelector(".legend");
  assert.ok(legend);
  const styleText = dom.window.document.querySelector("style")?.textContent || "";
  for (const cls of CANONICAL_STYLE_CONTRACT.requiredHighlightClasses) {
    assert.ok(styleText.includes(`.${cls.replace(" ", ".")}`));
  }
  for (const tableClass of CANONICAL_STYLE_CONTRACT.requiredTableClassNames) {
    assert.ok(dom.window.document.querySelector(`table.${tableClass}`));
  }
});

test("maximal coverage gate passes when conditions appear in main body", () => {
  const inventory = buildInventory();
  const html = buildMaximalFixtureHtml(inventory);
  ensureMaximalCoverage(inventory, html);
});

test("maximal coverage gate fails when a condition is appendix-only and reports sections", () => {
  const inventory = buildInventory();
  const missing = "Diabetic Ketoacidosis";
  const bullet = buildCoreBullet(missing);
  const row = buildCoverageRow(missing);
  const html = buildMaximalFixtureHtml(inventory)
    .replace(bullet, "")
    .replace(row, "");
  try {
    ensureMaximalCoverage(inventory, html);
    assert.fail("Expected coverage failure.");
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    assert.ok(message.includes("MAXIMAL_COVERAGE_FAILED"));
    assert.ok(Array.isArray((err as any).sections));
    assert.ok(((err as any).sections || []).length > 0);
    assert.strictEqual((err as any).mainBodyFound, true);
  }
});

test("renderStudyGuideHtml includes shared base CSS", () => {
  const stepA: StepAOutput = {
    lecture_title: "Test",
    slides: [],
    raw_facts: [],
    buckets: {
      dx: [],
      pathophys: [],
      clinical: [],
      labs: [],
      imaging: [],
      treatment: [],
      complications: [],
      risk_factors: [],
      epidemiology: [],
      red_flags: [],
      buzzwords: [],
    },
    discriminators: [],
    exam_atoms: [],
    abbrev_map: {},
    source_spans: [],
  };
  const stepB: StepBOutput = {
    high_yield_summary: [],
    rapid_approach_table: [],
    one_page_last_minute_review: [],
    compare_differential: [],
    quant_cutoffs: [],
    pitfalls: [],
    glossary: [],
    supplemental_glue: [],
  };
  const stepC: StepCOutput = {
    coverage_confidence: "Med",
    unparsed_items: [],
    omissions: [],
    conflicts: [],
    checks: {
      has_high_yield_summary: false,
      has_rapid_approach_table: false,
      has_one_page_review: false,
      slide_count_stepA: 0,
      slide_count_rendered: 0,
    },
  };
  const html = renderStudyGuideHtml({
    lectureTitle: "Test",
    buildUtc: "1970-01-01T00:00:00Z",
    stepA,
    stepB,
    stepC,
  });
  assert.ok(html.includes(BASE_STUDY_GUIDE_CSS));
});
