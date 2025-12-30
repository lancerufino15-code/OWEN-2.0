/// <reference types="node" />
/**
 * Tests machine study guide generation pipeline end-to-end.
 *
 * Used by: `npm test` to validate study guide extraction, synthesis, and rendering steps.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Step A chunking, JSON parsing/repair, payload assembly, and HTML truncation.
 *
 * Assumptions:
 * - Uses stubbed Responses API calls and an in-memory R2 bucket.
 */

import assert from "node:assert";
import test from "node:test";
import worker, {
  parseMachineSlideListFromTxt,
  parseMachineSlideBlocksFromTxt,
  applyStudyGuidePromptTemplate,
  buildStudyGuideStepAChunks,
  buildStudyGuideFilename,
  buildStudyGuideStoredKey,
  buildStudyGuidePayload,
  stripStudyGuideForbiddenKeys,
  parseStudyGuideJson,
  parseStudyGuideJsonWithRepair,
  assessStepAQuality,
  mergeStepAChunks,
  mergeStepAExtractAndDerived,
  STUDY_GUIDE_MODEL,
  withMaxOutputTokens,
  truncateAtHtmlEnd,
} from "./src/index";
import { STUDY_GUIDE_STEP_A_EXTRACT_PROMPT } from "./src/prompts/study_guide_stepA_extract";
import type { StepAOutput } from "./src/machine/render_study_guide_html";
import type { Env } from "./src/types";

/**
 * Minimal in-memory R2Bucket stub for study guide tests.
 */
class MemoryR2Bucket {
  private objects = new Map<string, { body: string; httpMetadata?: Record<string, string> }>();

  async put(key: string, value: string, opts?: { httpMetadata?: Record<string, string> }) {
    this.objects.set(key, { body: value, httpMetadata: opts?.httpMetadata });
  }

  async get(key: string) {
    const stored = this.objects.get(key);
    if (!stored) return null;
    return {
      body: stored.body,
      text: async () => stored.body,
      httpMetadata: stored.httpMetadata,
    };
  }

  async head(key: string) {
    const stored = this.objects.get(key);
    if (!stored) return null;
    return { key, size: stored.body.length, httpMetadata: stored.httpMetadata };
  }
}

/**
 * Build a minimal Step A chunk fixture for a given slide range.
 */
const buildStepAChunkJson = (start: number, end: number) => ({
  lecture_title: "Test Lecture",
  chunk: { start_slide: start, end_slide: end },
  slides: [
    {
      n: start,
      page: start,
      sections: [
        {
          heading: "General",
          facts: [
            { text: `Fact ${start}`, tags: ["disease"], numbers: [] },
          ],
        },
      ],
      tables: [],
    },
  ],
});

const STEP_A_DERIVED_JSON = {
  raw_facts: ["Fact"],
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
  exam_atoms: ["Atom"],
  abbrev_map: {},
  source_spans: [],
};

const FACT_REWRITE_JSON = { topics: [], spans: [] };

/**
 * Parse slide range markers from the Step A prompt text.
 *
 * @param prompt - Prompt string that includes slide range.
 * @returns Parsed start/end slide numbers with defaults.
 */
function parseChunkRange(prompt: string) {
  const match = prompt.match(/slides\s+(\d+)\s*-\s*(\d+)/i);
  if (!match) return { start: 1, end: 1 };
  return { start: Number(match[1]), end: Number(match[2]) };
}

/**
 * Create a Responses API stub that returns canned payloads by prompt type.
 *
 * @param opts - Hooks for truncation simulation and per-chunk callbacks.
 * @returns Fetch-compatible stub function.
 */
function createResponsesStub(opts?: {
  truncateThreshold?: number;
  onChunkCall?: (info: { start: number; end: number; length: number; truncated: boolean }) => void;
}) {
  const threshold = opts?.truncateThreshold ?? Number.POSITIVE_INFINITY;
  return async (input: any, init?: any) => {
    const request = input instanceof Request ? input : null;
    const url = typeof input === "string" ? input : request ? request.url : "";
    if (!url.endsWith("/responses")) {
      throw new Error(`Unexpected fetch: ${url}`);
    }
    const rawBody = request ? await request.text() : typeof init?.body === "string" ? init.body : "";
    const payload = rawBody ? JSON.parse(rawBody) : {};
    const prompt = payload?.input?.[0]?.content?.[0]?.text || "";
    const expectsJson = payload?.text?.format?.type === "json_object";
    if (!expectsJson) {
      return new Response(
        JSON.stringify({
          output_text: [
            "<html><head><style>body{font-family:Arial}</style></head><body><h1>Guide</h1></body></html>",
          ],
        }),
        { status: 200, headers: { "content-type": "application/json" } },
      );
    }
    if (prompt.includes("CHUNK_TEXT:")) {
      const { start, end } = parseChunkRange(prompt);
      const truncated = prompt.length > threshold;
      opts?.onChunkCall?.({ start, end, length: prompt.length, truncated });
      if (truncated) {
        return new Response(
          JSON.stringify({ output_text: [`{"lecture_title":"Test","chunk":{"start_slide":${start},"end_slide":${end}},"slides":[`] }),
          { status: 200, headers: { "content-type": "application/json" } },
        );
      }
      return new Response(
        JSON.stringify({ output_text: [JSON.stringify(buildStepAChunkJson(start, end))] }),
        { status: 200, headers: { "content-type": "application/json" } },
      );
    }
    if (prompt.includes("STEP_A1_JSON:")) {
      return new Response(
        JSON.stringify({ output_text: [JSON.stringify(STEP_A_DERIVED_JSON)] }),
        { status: 200, headers: { "content-type": "application/json" } },
      );
    }
    if (prompt.includes("FACT_REGISTRY_JSON")) {
      return new Response(
        JSON.stringify({ output_text: [JSON.stringify(FACT_REWRITE_JSON)] }),
        { status: 200, headers: { "content-type": "application/json" } },
      );
    }
    return new Response(
      JSON.stringify({ output_text: [JSON.stringify(STEP_A_DERIVED_JSON)] }),
      { status: 200, headers: { "content-type": "application/json" } },
    );
  };
}

test("parseMachineSlideListFromTxt returns sorted slide list and count", () => {
  const input = [
    "Slide 2 (p.2):",
    "Two",
    "",
    "Slide 1 (p.1):",
    "One",
    "",
    "Slide 3 (p.3):",
    "Three",
  ].join("\n");

  const parsed = parseMachineSlideListFromTxt(input);

  assert.strictEqual(parsed.slideCount, 3);
  assert.deepStrictEqual(parsed.slides, [
    { n: 1, title: "", page: 1 },
    { n: 2, title: "", page: 2 },
    { n: 3, title: "", page: 3 },
  ]);
});

test("applyStudyGuidePromptTemplate replaces placeholders only", () => {
  const template = "A {{BUILD_UTC_LITERAL}} B {{SLIDE_COUNT_LITERAL}} C {{JSON_OF_SLIDE_LIST}} D {{DOC_TEXT}} E";
  const output = applyStudyGuidePromptTemplate(template, {
    "{{BUILD_UTC_LITERAL}}": "1970-01-01T00:00:00Z",
    "{{SLIDE_COUNT_LITERAL}}": "3",
    "{{JSON_OF_SLIDE_LIST}}": "[{\"n\":1}]",
    "{{DOC_TEXT}}": "Slide 1 (p.1):",
  });
  assert.strictEqual(
    output,
    "A 1970-01-01T00:00:00Z B 3 C [{\"n\":1}] D Slide 1 (p.1): E",
  );
});

test("buildStudyGuideStoredKey is deterministic for same title/docId", () => {
  const keyA = buildStudyGuideStoredKey("doc123", "My Lecture");
  const keyB = buildStudyGuideStoredKey("doc123", "My Lecture");
  assert.strictEqual(keyA, keyB);
  assert.strictEqual(keyA, "machine/study-guides/My-Lecture/Study_Guide_My-Lecture.html");
  assert.strictEqual(buildStudyGuideFilename("My Lecture"), "Study_Guide_My Lecture.html");
});

test("buildStudyGuidePayload builds a minimal payload for gpt-4o", () => {
  const payload = buildStudyGuidePayload(STUDY_GUIDE_MODEL, "<html></html>", 4321) as Record<string, unknown>;
  assert.strictEqual(payload.model, STUDY_GUIDE_MODEL);
  assert.deepStrictEqual(payload.input, [
    {
      role: "user",
      content: [{ type: "input_text", text: "<html></html>" }],
    },
  ]);
  assert.strictEqual(payload.max_output_tokens, 4321);
  assert.strictEqual(payload.temperature, 0);
  assert.strictEqual(payload.top_p, 1);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(payload, "seed"), false);
  const keys = Object.keys(payload).sort();
  assert.deepStrictEqual(keys, ["input", "max_output_tokens", "model", "temperature", "top_p"]);
});

test("stripStudyGuideForbiddenKeys removes unsupported params", () => {
  const payload = buildStudyGuidePayload("gpt-4o", "<html></html>", 9999) as Record<string, unknown>;
  payload.stop = ["</html>"];
  payload.stop_sequences = ["</html>"];
  payload.stopSequences = ["</html>"];
  payload.seed = 42;
  payload.frequency_penalty = 1;
  payload.presence_penalty = 1;
  payload.max_completion_tokens = 123;
  payload.max_tokens = 456;
  payload.n = 1;
  payload.stream = false;
  payload.response = { format: "json_object" };
  const sanitized = stripStudyGuideForbiddenKeys(payload);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "stop"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "stop_sequences"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "stopSequences"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "temperature"), true);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "top_p"), true);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "seed"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "frequency_penalty"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "presence_penalty"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "max_completion_tokens"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "max_tokens"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "n"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "stream"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "response"), false);
});

test("study guide outbound payload omits seed", () => {
  const payload = buildStudyGuidePayload(STUDY_GUIDE_MODEL, "<html></html>", 1200) as Record<string, unknown>;
  const sanitized = stripStudyGuideForbiddenKeys(payload);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(sanitized, "seed"), false);
});

test("withMaxOutputTokens prefers max_output_tokens and removes legacy keys", () => {
  const payload = withMaxOutputTokens(
    { model: STUDY_GUIDE_MODEL, input: "<html></html>", max_tokens: 111, max_completion_tokens: 222 } as Record<string, unknown>,
    333,
  );
  assert.strictEqual(payload.max_output_tokens, 333);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(payload, "max_completion_tokens"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(payload, "max_tokens"), false);
});

test("truncateAtHtmlEnd returns truncated html", () => {
  const result = truncateAtHtmlEnd("<html>ok</html> trailing");
  if (!result.ok) {
    assert.fail(result.error);
  }
  assert.strictEqual(result.html, "<html>ok</html>");
});

test("truncateAtHtmlEnd errors when missing end tag", () => {
  const result = truncateAtHtmlEnd("<html>missing end");
  if (result.ok) {
    assert.fail("Expected missing </html> to error.");
  }
  assert.strictEqual(
    result.error,
    "Model output missing </html>. Study guide generation failed.",
  );
});

test("truncateAtHtmlEnd uses first end tag when multiple exist", () => {
  const result = truncateAtHtmlEnd("<html>first</html> extra </html>");
  if (!result.ok) {
    assert.fail(result.error);
  }
  assert.strictEqual(result.html, "<html>first</html>");
});

test("assessStepAQuality flags undersized Step A output", () => {
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
  const result = assessStepAQuality(stepA, [], 100);
  assert.strictEqual(result.ok, false);
});

test("parseStudyGuideJson accepts valid step payloads", () => {
  const stepAJson = JSON.stringify({
    lecture_title: "Test",
    chunk: { start_slide: 1, end_slide: 1 },
    slides: [
      {
        n: 1,
        page: 1,
        sections: [
          {
            heading: "General",
            facts: [
              {
                text: "Fact",
                tags: ["buzz"],
                numbers: [{ value: "5", unit: "mg" }],
              },
            ],
          },
        ],
        tables: [{ caption: "Table", headers: ["A"], rows: [["B"]] }],
      },
    ],
    raw_facts: ["Fact"],
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
    exam_atoms: ["Atom"],
    abbrev_map: { HF: "heart failure" },
    source_spans: [],
  });
  const stepBJson = JSON.stringify({
    high_yield_summary: ["Summary"],
    rapid_approach_table: [{ clue: "C", think_of: "T", why: "W", confirm: "K" }],
    one_page_last_minute_review: ["Review"],
    compare_differential: [{ topic: "Dx", rows: [{ dx1: "A", dx2: "B", how_to_tell: "C" }] }],
    quant_cutoffs: [{ item: "Item", value: "1", note: "Note" }],
    pitfalls: ["Pitfall"],
    glossary: [{ term: "Term", definition: "Def" }],
    supplemental_glue: [],
  });
  const stepCJson = JSON.stringify({
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
  });

  const parsedA = parseStudyGuideJson(stepAJson, "A") as any;
  const parsedB = parseStudyGuideJson(stepBJson, "B") as any;
  const parsedC = parseStudyGuideJson(stepCJson, "C") as any;

  assert.strictEqual(parsedA.lecture_title, "Test");
  assert.strictEqual(parsedB.high_yield_summary[0], "Summary");
  assert.strictEqual(parsedC.coverage_confidence, "Med");
});

test("buildStudyGuideStepAChunks splits 25 slides into five chunks", () => {
  const slides = Array.from({ length: 25 }, (_, i) => {
    const n = i + 1;
    return `Slide ${n} (p.${n}):\nContent ${n}`;
  }).join("\n\n");
  const parsed = parseMachineSlideBlocksFromTxt(slides);
  const chunks = buildStudyGuideStepAChunks(parsed.slides, 6, 12_000);
  assert.strictEqual(chunks.length, 5);
  assert.deepStrictEqual(
    chunks.map(chunk => [chunk.startSlide, chunk.endSlide]),
    [[1, 6], [7, 12], [13, 18], [19, 24], [25, 25]],
  );
});

test("mergeStepAChunks preserves order and de-duplicates", () => {
  const chunks = [
    {
      lecture_title: "Test",
      chunk: { start_slide: 1, end_slide: 2 },
      slides: [
        { n: 1, page: 1, sections: [], tables: [] },
        { n: 2, page: 2, sections: [], tables: [] },
      ],
    },
    {
      lecture_title: "Test",
      chunk: { start_slide: 3, end_slide: 4 },
      slides: [
        { n: 2, page: 2, sections: [], tables: [] },
        { n: 3, page: 3, sections: [], tables: [] },
      ],
    },
  ];
  const merged = mergeStepAChunks(chunks as any);
  assert.strictEqual(merged.slides.length, 3);
  assert.deepStrictEqual(merged.slides.map(slide => slide.n), [1, 2, 3]);
});

test("mergeStepAExtractAndDerived combines A1 and A2 outputs", () => {
  const extract = {
    lecture_title: "Test",
    slides: [{ n: 1, page: 1, sections: [], tables: [] }],
  };
  const derived = {
    raw_facts: ["Fact"],
    buckets: {
      dx: ["Dx"],
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
  const combined = mergeStepAExtractAndDerived(extract, derived);
  assert.strictEqual(combined.lecture_title, "Test");
  assert.strictEqual(combined.slides.length, 1);
  assert.strictEqual(combined.raw_facts[0], "Fact");
  assert.strictEqual(combined.buckets.dx[0], "Dx");
  assert.strictEqual(combined.exam_atoms[0], "Atom");
});

test("step A extract prompt avoids pipe tokens in tags example", () => {
  assert.strictEqual(STUDY_GUIDE_STEP_A_EXTRACT_PROMPT.includes("[\"disease\"|"), false);
  assert.ok(STUDY_GUIDE_STEP_A_EXTRACT_PROMPT.includes("\"tags\": [\"disease\", \"diagnostic\"]"));
  assert.strictEqual(STUDY_GUIDE_STEP_A_EXTRACT_PROMPT.includes("output '{}'"), false);
  assert.ok(STUDY_GUIDE_STEP_A_EXTRACT_PROMPT.includes("\"slides\": []"));
});

test("parseStudyGuideJsonWithRepair calls repair once on invalid JSON", async () => {
  let calls = 0;
  const repair = async () => {
    calls += 1;
    return JSON.stringify({
      lecture_title: "Test",
      chunk: { start_slide: 1, end_slide: 1 },
      slides: [],
    });
  };
  const parsed = await parseStudyGuideJsonWithRepair(
    "{\"lecture_title\":",
    "A1",
    repair,
  );
  assert.strictEqual(calls, 1);
  assert.strictEqual((parsed as any).lecture_title, "Test");
});

test("parseStudyGuideJsonWithRepair skips repair when JSON is valid", async () => {
  let calls = 0;
  const repair = async () => {
    calls += 1;
    return "{}";
  };
  const parsed = await parseStudyGuideJsonWithRepair(
    "{\"lecture_title\":\"Test\",\"chunk\":{\"start_slide\":1,\"end_slide\":1},\"slides\":[]}",
    "A1",
    repair,
  );
  assert.strictEqual(calls, 0);
  assert.strictEqual((parsed as any).lecture_title, "Test");
});

test("maximal study guide returns HTML and stored key", async () => {
  const originalFetch = globalThis.fetch;
  const bucket = new MemoryR2Bucket();
  let responsesCalls = 0;
  const stub = createResponsesStub();
  globalThis.fetch = async (input, init) => {
    responsesCalls += 1;
    return stub(input, init);
  };

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
      OWEN_INGEST: bucket,
      DOCS_KV: { put: async () => {} },
    } as unknown as Env;

    const txt = [
      "Slide 1 (p.1):",
      "Topic: Cardiology",
      `Details: ${"A".repeat(320)}`,
      "",
      "Slide 2 (p.2):",
      "Topic: Pulmonary",
      `Details: ${"B".repeat(320)}`,
    ].join("\n");

    const req = new Request("http://localhost/api/machine/generate-study-guide", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        mode: "maximal",
        lectureTitle: "Test Lecture",
        docId: "doc-test",
        txt,
      }),
    });

    const resp = await worker.fetch(req, env);
    assert.strictEqual(resp.status, 200);
    const body = await resp.json();
    const expectedKey = buildStudyGuideStoredKey("doc-test", "Test Lecture");
    assert.strictEqual(body.ok, true);
    assert.strictEqual(body.storedKey, expectedKey);
    assert.ok(body.sourceStoredKey);
    assert.ok(typeof body.downloadUrl === "string" && body.downloadUrl.includes(encodeURIComponent(expectedKey)));
    assert.ok(responsesCalls > 0);
    const stored = await bucket.get(expectedKey);
    assert.ok(stored);
    const storedText = stored ? await stored.text() : "";
    assert.ok(storedText.includes("</html>"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("step A truncation splits and succeeds", async () => {
  const originalFetch = globalThis.fetch;
  const bucket = new MemoryR2Bucket();
  const chunkCalls: Array<{ start: number; end: number; truncated: boolean }> = [];
  const stub = createResponsesStub({
    truncateThreshold: 8000,
    onChunkCall: (info) => {
      chunkCalls.push({ start: info.start, end: info.end, truncated: info.truncated });
    },
  });
  globalThis.fetch = async (input, init) => stub(input, init);

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
      OWEN_INGEST: bucket,
      DOCS_KV: { put: async () => {} },
    } as unknown as Env;

    const slideText = "A".repeat(1500);
    const txt = Array.from({ length: 6 }, (_, i) => `Slide ${i + 1} (p.${i + 1}):\n${slideText}`).join("\n\n");
    const req = new Request("http://localhost/api/machine/generate-study-guide", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        mode: "maximal",
        lectureTitle: "Split Lecture",
        docId: "doc-split",
        txt,
      }),
    });

    const resp = await worker.fetch(req, env);
    assert.strictEqual(resp.status, 200);
    const body = await resp.json();
    assert.strictEqual(body.ok, true);
    assert.ok(chunkCalls.some(call => call.truncated));
    assert.ok(chunkCalls.some(call => !call.truncated && call.end - call.start < 6));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("step A checkpoint resume reuses cached chunks", async () => {
  const originalFetch = globalThis.fetch;
  const originalNow = Date.now;
  const bucket = new MemoryR2Bucket();
  const chunkCalls = new Map<string, number>();
  let timeBudgetExceeded = false;
  let allowBudgetTrip = true;
  const baseTime = originalNow();
  const stub = createResponsesStub({
    onChunkCall: (info) => {
      const key = `${info.start}-${info.end}`;
      chunkCalls.set(key, (chunkCalls.get(key) || 0) + 1);
      if (allowBudgetTrip) {
        timeBudgetExceeded = true;
      }
    },
  });
  globalThis.fetch = async (input, init) => stub(input, init);
  Date.now = () => (timeBudgetExceeded ? baseTime + 30_000 : baseTime);

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
      OWEN_INGEST: bucket,
      DOCS_KV: { put: async () => {} },
    } as unknown as Env;

    const txt = Array.from({ length: 8 }, (_, i) => `Slide ${i + 1} (p.${i + 1}):\n${"B".repeat(200)}`).join("\n\n");
    const req = new Request("http://localhost/api/machine/generate-study-guide", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        mode: "maximal",
        lectureTitle: "Resume Lecture",
        docId: "doc-resume",
        txt,
      }),
    });

    const first = await worker.fetch(req, env);
    const firstBody = await first.json();
    assert.strictEqual(firstBody.ok, true);
    assert.strictEqual(firstBody.partial, true);

    const firstCount = chunkCalls.get("1-6") || 0;
    timeBudgetExceeded = false;
    allowBudgetTrip = false;

    const second = await worker.fetch(req, env);
    const secondBody = await second.json();
    assert.strictEqual(secondBody.ok, true);
    const secondCount = chunkCalls.get("1-6") || 0;
    assert.strictEqual(secondCount, firstCount);
  } finally {
    globalThis.fetch = originalFetch;
    Date.now = originalNow;
  }
});

test("generate-study-guide accepts legacy txt JSON", async () => {
  const originalFetch = globalThis.fetch;
  const bucket = new MemoryR2Bucket();
  const stub = createResponsesStub();
  globalThis.fetch = async (input, init) => stub(input, init);

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
      OWEN_INGEST: bucket,
      DOCS_KV: { put: async () => {} },
    } as unknown as Env;

    const txt = `Slide 1 (p.1):\n${"C".repeat(520)}`;
    const req = new Request("http://localhost/api/machine/generate-study-guide", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        mode: "canonical",
        lectureTitle: "Legacy Lecture",
        docId: "doc-legacy",
        txt,
      }),
    });

    const resp = await worker.fetch(req, env);
    const body = await resp.json();
    assert.strictEqual(body.ok, true);
    assert.ok(body.sourceStoredKey);
    const stored = await bucket.get(body.sourceStoredKey);
    assert.ok(stored);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("generate-study-guide accepts storedKey JSON", async () => {
  const originalFetch = globalThis.fetch;
  const bucket = new MemoryR2Bucket();
  const stub = createResponsesStub();
  globalThis.fetch = async (input, init) => stub(input, init);

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
      OWEN_INGEST: bucket,
      DOCS_KV: { put: async () => {} },
    } as unknown as Env;

    const txt = `Slide 1 (p.1):\n${"D".repeat(520)}`;
    const storedKey = "machine/txt/test-doc.txt";
    await bucket.put(storedKey, txt, { httpMetadata: { contentType: "text/plain; charset=utf-8" } });

    const req = new Request("http://localhost/api/machine/generate-study-guide", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        mode: "canonical",
        storedKey,
        filename: "Stored Lecture.txt",
        lectureTitle: "Stored Lecture",
        docId: "doc-stored",
      }),
    });

    const resp = await worker.fetch(req, env);
    const body = await resp.json();
    assert.strictEqual(body.ok, true);
    assert.ok(body.sourceStoredKey);
    const stored = await bucket.get(body.sourceStoredKey);
    assert.ok(stored);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("generate-study-guide accepts formData txtFile", async () => {
  const originalFetch = globalThis.fetch;
  const bucket = new MemoryR2Bucket();
  const stub = createResponsesStub();
  globalThis.fetch = async (input, init) => stub(input, init);

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
      OWEN_INGEST: bucket,
      DOCS_KV: { put: async () => {} },
    } as unknown as Env;

    const txt = `Slide 1 (p.1):\n${"E".repeat(520)}`;
    const form = new (globalThis as any).FormData();
    const blob = new (globalThis as any).Blob([txt], { type: "text/plain" });
    form.append("txtFile", blob, "Form Lecture.txt");
    form.append("mode", "canonical");

    const req = new Request("http://localhost/api/machine/generate-study-guide", {
      method: "POST",
      body: form,
    });

    const resp = await worker.fetch(req, env);
    const body = await resp.json();
    assert.strictEqual(body.ok, true);
    assert.ok(body.sourceStoredKey);
    const stored = await bucket.get(body.sourceStoredKey);
    assert.ok(stored);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
