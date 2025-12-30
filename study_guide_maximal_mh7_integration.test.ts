/// <reference types="node" />
/**
 * Integration tests for maximal MH7 study guide generation.
 *
 * Used by: `npm test` to validate the maximal study guide pipeline.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Fact registry rewrite, MH7 filtering, and HTML rendering with no placeholders.
 *
 * Assumptions:
 * - Responses API calls are stubbed; HTML validation uses JSDOM only for checks.
 */

import assert from "node:assert";
import test from "node:test";
import { JSDOM } from "jsdom";
import {
  applyStudyGuidePromptTemplate,
  callStudyGuideResponses,
  mergeStepAChunks,
  mergeStepAExtractAndDerived,
  parseStudyGuideJsonWithRepair,
} from "./src/index";
import type { Env } from "./src/types";
import type { StepAOutput } from "./src/machine/render_study_guide_html";
import { STUDY_GUIDE_MAXIMAL_FACT_REWRITE_PROMPT } from "./src/prompts/study_guide_maximal_fact_rewrite";
import { extractTopicInventoryFromSlides } from "./src/study_guides/inventory";
import {
  buildFactRegistryFromStepA,
  coerceFactRegistryRewrite,
  filterMh7Topics,
  type FactRegistry,
} from "./src/study_guides/fact_registry";
import { renderMaximalStudyGuideHtml } from "./src/study_guides/render_maximal_html";

type StepAChunkOutput = {
  lecture_title: string;
  chunk: { start_slide: number; end_slide: number };
  slides: StepAOutput["slides"];
};

test("maximal integration yields MH7 tables and no placeholders", async () => {
  const slides = [
    {
      n: 1,
      text: [
        "Tacrolimus",
        "Calcineurin inhibitors",
        "Hyperacute rejection",
        "Calcineurin -> NFAT -> IL-2 axis",
      ].join("\n"),
    },
  ];

  const stepAChunk: StepAChunkOutput = {
    lecture_title: "Transplant Pharm",
    chunk: { start_slide: 1, end_slide: 1 },
    slides: [
      {
        n: 1,
        page: 1,
        sections: [
          {
            heading: "Tacrolimus",
            facts: [
              { text: "Tacrolimus inhibits calcineurin to reduce IL-2 transcription", tags: ["treatment"], numbers: [] },
              { text: "Tacrolimus causes nephrotoxicity", tags: ["treatment"], numbers: [] },
              { text: "Tacrolimus causes neurotoxicity", tags: ["treatment"], numbers: [] },
              { text: "Tacrolimus is metabolized by CYP3A", tags: ["enzyme"], numbers: [] },
              { text: "Monitor tacrolimus trough levels", tags: ["lab"], numbers: [] },
              { text: "Tacrolimus used for transplant maintenance", tags: ["treatment"], numbers: [] },
            ],
          },
          {
            heading: "Calcineurin inhibitors",
            facts: [
              { text: "Calcineurin inhibitors block T-cell activation", tags: ["treatment"], numbers: [] },
              { text: "Calcineurin inhibitors cause nephrotoxicity", tags: ["treatment"], numbers: [] },
              { text: "Calcineurin inhibitors are metabolized by CYP3A", tags: ["enzyme"], numbers: [] },
              { text: "Calcineurin inhibitors used for maintenance immunosuppression", tags: ["treatment"], numbers: [] },
            ],
          },
          {
            heading: "Hyperacute rejection",
            facts: [
              { text: "Hyperacute rejection occurs within minutes after transplant", tags: ["disease"], numbers: [] },
              { text: "Hyperacute rejection is a B-cell antibody response", tags: ["disease"], numbers: [] },
              { text: "Treat hyperacute rejection with immediate graft removal", tags: ["treatment"], numbers: [] },
            ],
          },
        ],
        tables: [],
      },
    ],
  };

  const stepADerived = {
    raw_facts: [
      "Tacrolimus inhibits calcineurin to reduce IL-2 transcription",
      "Tacrolimus causes nephrotoxicity",
      "Tacrolimus causes neurotoxicity",
      "Tacrolimus is metabolized by CYP3A",
      "Monitor tacrolimus trough levels",
      "Tacrolimus used for transplant maintenance",
      "Calcineurin inhibitors block T-cell activation",
      "Calcineurin inhibitors cause nephrotoxicity",
      "Calcineurin inhibitors are metabolized by CYP3A",
      "Calcineurin inhibitors used for maintenance immunosuppression",
      "Hyperacute rejection occurs within minutes after transplant",
      "Hyperacute rejection is a B-cell antibody response",
      "Treat hyperacute rejection with immediate graft removal",
      "Calcineurin -> NFAT -> IL-2 axis",
    ],
    buckets: {
      dx: ["Hyperacute rejection"],
      pathophys: ["Calcineurin -> NFAT -> IL-2 axis"],
      clinical: [],
      labs: ["Tacrolimus trough"],
      imaging: [],
      treatment: ["Tacrolimus"],
      complications: [],
      risk_factors: [],
      epidemiology: [],
      red_flags: [],
      buzzwords: [],
    },
    discriminators: [],
    exam_atoms: [
      "Tacrolimus nephrotoxicity",
      "Tacrolimus neurotoxicity",
      "Hyperacute rejection within minutes",
    ],
    abbrev_map: {},
    source_spans: [
      { text: "Tacrolimus inhibits calcineurin to reduce IL-2 transcription", slides: [1], pages: [1] },
      { text: "Tacrolimus causes nephrotoxicity", slides: [1], pages: [1] },
      { text: "Tacrolimus causes neurotoxicity", slides: [1], pages: [1] },
      { text: "Tacrolimus is metabolized by CYP3A", slides: [1], pages: [1] },
      { text: "Tacrolimus used for transplant maintenance", slides: [1], pages: [1] },
    ],
  };

  const env = {
    OPENAI_API_KEY: "test-key",
    OPENAI_API_BASE: "https://api.openai.com/v1",
  } as Env;

  let rewritePayload = "";
  let callCount = 0;
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => {
    callCount += 1;
    const payload = callCount === 1 ? JSON.stringify(stepAChunk) : rewritePayload;
    return new Response(JSON.stringify({ output_text: [payload] }), {
      status: 200,
      headers: { "content-type": "application/json" },
    });
  };

  try {
    const stepARaw = await callStudyGuideResponses(
      env,
      "req-1",
      "A1",
      "gpt-4o",
      "step A prompt",
      2000,
      { expectsJson: true },
    );
    const stepAChunkParsed = JSON.parse(stepARaw) as StepAChunkOutput;
    const stepAExtract = mergeStepAChunks([stepAChunkParsed]);
    const stepA = mergeStepAExtractAndDerived(stepAExtract, stepADerived);

    const inventory = extractTopicInventoryFromSlides(slides);
    const registry = buildFactRegistryFromStepA({
      stepA,
      inventory,
      slides: slides.map(slide => ({ n: slide.n, text: slide.text, page: 1 })),
    });

    const rewriteJson = { topics: registry.topics };
    rewritePayload = JSON.stringify(rewriteJson);
    const rewritePrompt = applyStudyGuidePromptTemplate(STUDY_GUIDE_MAXIMAL_FACT_REWRITE_PROMPT, {
      "{{FACT_REGISTRY_JSON}}": JSON.stringify(registry),
    });
    const rewriteRaw = await callStudyGuideResponses(
      env,
      "req-1",
      "maximal-facts",
      "gpt-4o",
      rewritePrompt,
      2000,
      { expectsJson: true },
    );
    const rewritten = await parseStudyGuideJsonWithRepair<FactRegistry>(rewriteRaw, "maximal-facts", async () => {
      return JSON.stringify(rewriteJson);
    });
    const rewrittenRegistry = coerceFactRegistryRewrite(registry, rewritten);
    const mh7 = filterMh7Topics(rewrittenRegistry.topics, { requireDrugCoverage: true });
    const filteredRegistry = { ...rewrittenRegistry, topics: mh7.kept };

    const html = renderMaximalStudyGuideHtml({
      lectureTitle: stepA.lecture_title,
      buildUtc: "1970-01-01T00:00:00Z",
      slideCount: slides.length,
      slides,
      inventory,
      registry: filteredRegistry,
      mh7: { omitted: mh7.omitted, minFacts: mh7.minFacts },
      stepA,
    });

    assert.ok(!html.includes("Not stated in lecture"));
    assert.ok(!html.includes("N/A"));
    const dom = new JSDOM(html);
    const rapidHeaders = Array.from(dom.window.document.querySelectorAll("#rapid-approach-summary th")).map(th =>
      (th.textContent || "").trim(),
    );
    assert.ok(rapidHeaders.includes("Why (discriminator)"));

    const compareTables = Array.from(dom.window.document.querySelectorAll("#treatments-management table.compare"));
    const drugClassTable = compareTables.find(table => (table.textContent || "").includes("Best use"));
    assert.ok(drugClassTable);
    const drugHeaders = Array.from(drugClassTable!.querySelectorAll("th")).map(th => (th.textContent || "").trim());
    assert.ok(drugHeaders.includes("Why (discriminator)"));

    assert.strictEqual(callCount, 2);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
