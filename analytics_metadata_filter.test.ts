/// <reference types="node" />
/**
 * Tests metadata/token normalization helpers used for analytics tagging.
 *
 * Used by: `npm test` (Node test runner) to validate prompt-cleaning behavior.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Stopword removal, token normalization, alias grouping, and canonical key building.
 *
 * Assumptions:
 * - Stopword/preposition lists are English-centric and stable across runs.
 */

import assert from "node:assert";
import test from "node:test";
import {
  buildMetaWeights,
  buildCanonicalPromptKey,
  buildPromptAliases,
  mergeAliasList,
  cleanPromptPreserveOrder,
  cleanPromptToSentence,
  extractNonStopMeta,
  filterMetadata,
  removeArticlesAndPrepositionsFromString,
  zeroOutStopMeta,
  type MetaWeights,
} from "./analytics_metadata_filter";

test("buildMetaWeights counts, trims, and collapses whitespace without mutating input", () => {
  const values = [" the ", "patient", "patient", "  ", "IN   spite   OF", "Vis-à-vis"];
  const snapshot = values.slice();

  const weights = buildMetaWeights(values);

  assert.deepStrictEqual(values, snapshot);
  assert.deepStrictEqual(weights, { patient: 2 });
});

test("zeroOutStopMeta flags stopwords and keeps non-stops unchanged", () => {
  const weights: MetaWeights = {
    the: 2,
    "in spite of": 1,
    "vis-à-vis": 1,
    hypertension: 3,
    patient: 1,
    Amongst: 4,
  };
  const original = { ...weights };

  const zeroed = zeroOutStopMeta(weights);

  assert.strictEqual(zeroed.the, 0);
  assert.strictEqual(zeroed["in spite of"], 0);
  assert.strictEqual(zeroed["vis-à-vis"], 0);
  assert.strictEqual(zeroed.Amongst, 0);
  assert.strictEqual(zeroed.hypertension, 3);
  assert.strictEqual(zeroed.patient, 1);

  assert.deepStrictEqual(weights, original);
  assert.notStrictEqual(zeroed, weights);
});

test("extractNonStopMeta returns sorted non-stop keys capped by limit", () => {
  const weights: MetaWeights = {
    alpha: 2,
    zeta: 2,
    patient: 1,
    hypertension: 3,
  };

  const result = extractNonStopMeta(weights, 3);

  assert.deepStrictEqual(result, ["hypertension", "alpha", "zeta"]);
});

test("filterMetadata pipeline zeroes stops and returns only non-stop strings", () => {
  const values = ["the", "in spite of", "hypertension", "patient", "patient", "vis-à-vis", "Amongst"];

  const result = filterMetadata(values, 10);

  assert.deepStrictEqual(result, ["hypertension", "patient", "patient"]);
});

test("filterMetadata tokenizes prompts and drops stopwords like about/within/the", () => {
  const values = [
    "tell me about different casts",
    "tell me the difference in lab values and the casts within urine",
  ];

  const result = filterMetadata(values, 10);

  assert.deepStrictEqual(result, ["casts", "lab", "values", "casts", "urine"]);
  assert.ok(!result.includes("about"));
  assert.ok(!result.includes("within"));
  assert.ok(!result.includes("the"));
});

test("cleanPromptToSentence preserves order and joins hepatitis qualifiers", () => {
  const input = "Tell me about the best drugs for people with EBV and Hepatitis B";
  const result = cleanPromptToSentence(input);
  assert.strictEqual(result.cleaned, "drugs EBV Hepatitis B");
  assert.deepStrictEqual(result.topics, [
    "drugs EBV Hepatitis B",
    "drugs EBV",
    "EBV Hepatitis B",
    "drugs",
    "EBV",
  ]);
});

test("cleanPromptPreserveOrder keeps left-to-right order", () => {
  const input = "Tell me about the diseases in this lecture, their clinical features and their relevant lab values";
  const cleaned = cleanPromptPreserveOrder(input);
  assert.strictEqual(typeof cleaned, "string");
  assert.strictEqual(cleaned, "diseases lecture, their clinical features their relevant lab values");
});

test("buildCanonicalPromptKey groups reordered phrasing into one bucket", () => {
  const inputA = "clinical pathophysiology methanol ethylene glycol poisoning?";
  const inputB = "ethylene glycol methanol poisoning pathophysiology clinical";
  const cleanedA = cleanPromptPreserveOrder(inputA);
  const cleanedB = cleanPromptPreserveOrder(inputB);
  const keyA = buildCanonicalPromptKey(cleanedA);
  const keyB = buildCanonicalPromptKey(cleanedB);
  assert.strictEqual(keyA, keyB);
  const counts = new Map<string, number>();
  counts.set(keyA, (counts.get(keyA) || 0) + 1);
  counts.set(keyB, (counts.get(keyB) || 0) + 1);
  assert.strictEqual(counts.size, 1);
  assert.strictEqual(counts.get(keyA), 2);
});

test("buildPromptAliases caps to five and preserves order", () => {
  const cleaned = "clinical pathophysiology methanol ethylene glycol poisoning";
  const aliases = buildPromptAliases(cleaned, 5);
  assert.deepStrictEqual(aliases, [
    "clinical pathophysiology methanol",
    "clinical pathophysiology",
    "pathophysiology methanol ethylene",
    "pathophysiology methanol",
    "methanol ethylene glycol",
  ]);
});

test("mergeAliasList caps to five and preserves insertion order", () => {
  const existing = ["alpha beta", "beta gamma", "gamma delta"];
  const incoming = ["alpha beta", "delta epsilon", "epsilon zeta", "zeta eta"];
  const merged = mergeAliasList(existing, incoming, 5);
  assert.deepStrictEqual(merged, [
    "alpha beta",
    "beta gamma",
    "gamma delta",
    "delta epsilon",
    "epsilon zeta",
  ]);
});

test("removeArticlesAndPrepositionsFromString strips inline stopwords and keeps order", () => {
  const input = "tell me about different casts within urine";
  const cleaned = removeArticlesAndPrepositionsFromString(input);
  assert.strictEqual(cleaned, "tell me different casts urine");
});

test("removeArticlesAndPrepositionsFromString drops multiword prepositions", () => {
  const input = "consider it vis-à-vis in spite of the plan";
  const cleaned = removeArticlesAndPrepositionsFromString(input);
  assert.strictEqual(cleaned, "consider it plan");
});

test("removeArticlesAndPrepositionsFromString removes stopwords inside sentences", () => {
  const input = "tell me about diseases in this lecture";
  const cleaned = removeArticlesAndPrepositionsFromString(input);
  assert.strictEqual(cleaned, "tell me diseases lecture");
});

test("removeArticlesAndPrepositionsFromString handles multiword and case-insensitive stops", () => {
  const input = "in spite of the findings";
  const cleaned = removeArticlesAndPrepositionsFromString(input);
  assert.strictEqual(cleaned, "findings");
});

test("removeArticlesAndPrepositionsFromString strips mixed-case articles/prepositions", () => {
  const input = "Tell Me About Casts";
  const cleaned = removeArticlesAndPrepositionsFromString(input);
  assert.strictEqual(cleaned, "tell me casts");
});
