/// <reference types="node" />
/**
 * Tests Step B validator rules for study guide synthesis outputs.
 *
 * Used by: `npm test` to validate the Step B quality gates.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Bullet length, redundancy, coverage, and synthesis minimums.
 *
 * Assumptions:
 * - Step B outputs are validated in isolation from runtime prompt execution.
 */

import assert from "node:assert";
import test from "node:test";
import { validateStepB, validateSynthesis, validateSynthesisOrThrow } from "./src/machine/study_guide_stepB_validator";
import type { StepAOutput, StepBOutput } from "./src/machine/render_study_guide_html";

const atoms = [
  "alpha",
  "bravo",
  "charlie",
  "delta",
  "echo",
  "foxtrot",
  "golf",
  "hotel",
  "india",
  "juliet",
  "kilo",
  "lima",
];

/**
 * Build a minimal Step A payload for validator tests.
 */
function buildStepABase(): StepAOutput {
  return {
    lecture_title: "Test",
    slides: [],
    raw_facts: ["alpha fact", "bravo fact"],
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
    exam_atoms: atoms,
    abbrev_map: { HF: "heart failure" },
    source_spans: [],
  };
}

/**
 * Build a valid Step B payload that should pass baseline checks.
 */
function buildValidStepB(): StepBOutput {
  return {
    high_yield_summary: atoms.slice(0, 8).map(atom => `HY ${atom} sign`),
    rapid_approach_table: Array.from({ length: 10 }, (_, i) => {
      const idx = i + 1;
      return {
        clue: `Clue ${idx}`,
        think_of: `Dx${idx}`,
        why: `Why ${idx} detail`,
        confirm: `Test ${idx}`,
      };
    }),
    one_page_last_minute_review: atoms.map(atom => `Review ${atom} cue`),
    compare_differential: [
      {
        topic: "Topic A",
        rows: Array.from({ length: 4 }, (_, i) => ({
          dx1: `A${i + 1}`,
          dx2: `B${i + 1}`,
          how_to_tell: `Tell A${i + 1} vs B${i + 1}`,
        })),
      },
      {
        topic: "Topic B",
        rows: Array.from({ length: 4 }, (_, i) => ({
          dx1: `C${i + 1}`,
          dx2: `D${i + 1}`,
          how_to_tell: `Tell C${i + 1} vs D${i + 1}`,
        })),
      },
    ],
    quant_cutoffs: [],
    pitfalls: [],
    glossary: [],
    supplemental_glue: [],
  };
}

test("validateStepB flags bullets that exceed word limits", () => {
  const stepA = buildStepABase();
  const stepB = buildValidStepB();
  stepB.high_yield_summary[0] = "word ".repeat(20).trim();
  const failures = validateStepB(stepA, stepB);
  const codes = failures.map(item => item.code);
  assert.strictEqual(codes.includes("BULLET_TOO_LONG"), true);
});

test("validateStepB flags too few bullets", () => {
  const stepA = buildStepABase();
  const stepB = buildValidStepB();
  stepB.high_yield_summary = ["Short one", "Short two"];
  const failures = validateStepB(stepA, stepB);
  const codes = failures.map(item => item.code);
  assert.strictEqual(codes.includes("TOO_FEW_BULLETS"), true);
});

test("validateStepB flags redundancy and n-gram overlap", () => {
  const stepA = buildStepABase();
  const stepB = buildValidStepB();
  stepB.one_page_last_minute_review = Array.from({ length: 12 }, () => "Repeatable token line");
  const failures = validateStepB(stepA, stepB);
  const codes = failures.map(item => item.code);
  assert.strictEqual(codes.includes("REDUNDANT_BULLETS"), true);
  assert.strictEqual(codes.includes("HIGH_NGRAM_OVERLAP"), true);
});

test("validateStepB flags low coverage", () => {
  const stepA = buildStepABase();
  const stepB = buildValidStepB();
  stepB.high_yield_summary = Array.from({ length: 8 }, (_, i) => `Unrelated ${i + 1}`);
  stepB.one_page_last_minute_review = Array.from({ length: 12 }, (_, i) => `Offtopic ${i + 1}`);
  const failures = validateStepB(stepA, stepB);
  const codes = failures.map(item => item.code);
  assert.strictEqual(codes.includes("LOW_COVERAGE"), true);
});

test("validateStepB flags supplemental glue violations", () => {
  const stepA = buildStepABase();
  const stepB = buildValidStepB();
  stepB.supplemental_glue = ["Unrelated glue content"];
  const failures = validateStepB(stepA, stepB);
  const codes = failures.map(item => item.code);
  assert.strictEqual(codes.includes("GLUE_RULE_VIOLATION"), true);
});

test("validateSynthesis flags empty synthesis sections", () => {
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
  const failures = validateSynthesis(stepB);
  const codes = failures.map(item => item.code);
  assert.strictEqual(codes.includes("SYNTHESIS_TOO_FEW"), true);
});

test("validateSynthesisOrThrow throws on below-minimum counts", () => {
  const stepB: StepBOutput = {
    high_yield_summary: ["One"],
    rapid_approach_table: Array.from({ length: 2 }, () => ({
      clue: "",
      think_of: "",
      why: "",
      confirm: "",
    })),
    one_page_last_minute_review: ["Two"],
    compare_differential: [],
    quant_cutoffs: [],
    pitfalls: [],
    glossary: [],
    supplemental_glue: [],
  };
  assert.throws(() => validateSynthesisOrThrow(stepB));
});
