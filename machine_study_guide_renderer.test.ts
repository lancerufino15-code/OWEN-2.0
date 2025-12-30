/// <reference types="node" />
/**
 * Tests study guide HTML renderer output shape.
 *
 * Used by: `npm test` to validate HTML structure for rendered study guides.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Required headings, HTML termination, and highlight classes.
 *
 * Assumptions:
 * - Renderer is pure (string output) and does not require DOM globals.
 */

import assert from "node:assert";
import test from "node:test";
import { renderStudyGuideHtml, type StepAOutput, type StepBOutput, type StepCOutput } from "./src/machine/render_study_guide_html";

test("renderStudyGuideHtml outputs required headings", () => {
  const stepA: StepAOutput = {
    lecture_title: "Test Lecture",
    slides: [
      {
        n: 1,
        page: 1,
        sections: [
          {
            heading: "General",
            facts: [
              { text: "Fact", tags: ["buzz"], numbers: [{ value: "5", unit: "mg" }] },
            ],
          },
        ],
        tables: [{ caption: "Table", headers: ["A"], rows: [["B"]] }],
      },
    ],
    raw_facts: ["Fact"],
    buckets: {
      dx: ["Asthma"],
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
    exam_atoms: ["Atom"],
    abbrev_map: { HF: "heart failure" },
    source_spans: [],
  };

  const stepB: StepBOutput = {
    high_yield_summary: ["Asthma baseline summary"],
    rapid_approach_table: [{ clue: "C", think_of: "T", why: "W", confirm: "K" }],
    one_page_last_minute_review: ["Review"],
    compare_differential: [{ topic: "Dx", rows: [{ dx1: "A", dx2: "B", how_to_tell: "C" }] }],
    quant_cutoffs: [{ item: "Item", value: "1", note: "Note" }],
    pitfalls: ["Pitfall"],
    glossary: [{ term: "Term", definition: "Def" }],
    supplemental_glue: [],
  };

  const stepC: StepCOutput = {
    coverage_confidence: "Med",
    unparsed_items: [{ slide: 1, note: "Note" }],
    omissions: [{ slide: 1, note: "Miss" }],
    conflicts: [{ slide: 1, note: "Conflict" }],
    checks: {
      has_high_yield_summary: true,
      has_rapid_approach_table: true,
      has_one_page_review: true,
      slide_count_stepA: 1,
      slide_count_rendered: 1,
    },
  };

  const html = renderStudyGuideHtml({
    lectureTitle: "Test Lecture",
    buildUtc: "1970-01-01T00:00:00Z",
    stepA,
    stepB,
    stepC,
    sourceTextBySlide: { "1": "Slide text" },
  });

  const requiredHeadings = [
    "Output Identity",
    "Document Map",
    "High-Yield Summary",
    "Rapid-Approach Summary (Global)",
    "One-Page Last-Minute Review",
    "Slide-by-Slide Appendix",
    "Source Note &amp; Quality Assurance",
  ];

  for (const heading of requiredHeadings) {
    assert.strictEqual(html.includes(heading), true);
  }
  assert.strictEqual(html.endsWith("</html>"), true);
  assert.strictEqual(html.includes(".hl.disease"), true);
  assert.strictEqual(html.includes("<span class=\"hl buzz\">"), true);
});
