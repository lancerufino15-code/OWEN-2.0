/// <reference types="node" />
/**
 * Tests JSON extraction/repair helpers for study guide outputs.
 *
 * Used by: `npm test` to validate JSON repair and retry logic.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Extract-first-object parsing, escape fixing, and truncation retries.
 *
 * Assumptions:
 * - Responses API calls are stubbed; no external network access.
 */

import assert from "node:assert";
import test from "node:test";
import {
  callStudyGuideResponses,
  extractFirstJsonObject,
  parseStudyGuideJsonWithRepair,
  parseStudyGuideJsonWithRepairAndRetry,
  repairJsonMinimal,
  STUDY_GUIDE_MODEL,
} from "./src/index";
import type { Env } from "./src/types";

test("extractFirstJsonObject returns clean JSON", () => {
  const raw = "{\"a\":1}";
  assert.strictEqual(extractFirstJsonObject(raw), raw);
});

test("extractFirstJsonObject ignores trailing suffix", () => {
  const raw = "{\"a\":[1,2]} trailing text";
  assert.strictEqual(extractFirstJsonObject(raw), "{\"a\":[1,2]}");
});

test("extractFirstJsonObject drops duplicated closers suffix", () => {
  const raw = "{\"a\":[{\"b\":1}]} ] }";
  assert.strictEqual(extractFirstJsonObject(raw), "{\"a\":[{\"b\":1}]}");
});

test("repairJsonMinimal trims trailing prose after JSON", () => {
  const raw = "{\"a\":1} trailing prose";
  const minimal = repairJsonMinimal(raw);
  assert.ok(minimal);
  assert.strictEqual(minimal?.repaired, "{\"a\":1}");
});

test("parseStudyGuideJsonWithRepair parses first object before suffix", async () => {
  let repairCalls = 0;
  const raw = "{\"a\":[{\"b\":1}]} ] }";
  const parsed = await parseStudyGuideJsonWithRepair<{ a: Array<{ b: number }> }>(
    raw,
    "A1",
    async () => {
      repairCalls += 1;
      return "{\"a\":[]}";
    },
  );
  assert.strictEqual(repairCalls, 0);
  assert.strictEqual(parsed.a[0].b, 1);
});

test("parseStudyGuideJsonWithRepairAndRetry retries on truncation", async () => {
  let retryCalls = 0;
  const result = await parseStudyGuideJsonWithRepairAndRetry<{ a: number }>({
    raw: "{\"a\":1",
    label: "A8",
    retry: async () => {
      retryCalls += 1;
      return "{\"a\":3}";
    },
    maxTruncationRetries: 1,
  });

  assert.strictEqual(retryCalls, 1);
  assert.strictEqual(result.a, 3);
});

test("parseStudyGuideJsonWithRepairAndRetry falls back after truncation retries", async () => {
  let retryCalls = 0;
  const result = await parseStudyGuideJsonWithRepairAndRetry<{
    lecture_title: string;
    chunk: { start_slide: number; end_slide: number };
    slides: unknown[];
  }>({
    raw: "{\"a\":1",
    label: "A9",
    retry: async () => {
      retryCalls += 1;
      return "{\"a\":1";
    },
    fallback: () => ({ lecture_title: "", chunk: { start_slide: 1, end_slide: 1 }, slides: [] }),
    maxTruncationRetries: 1,
  });

  assert.strictEqual(retryCalls, 1);
  assert.strictEqual(result.lecture_title, "");
  assert.deepStrictEqual(result.slides, []);
});

test("parseStudyGuideJsonWithRepairAndRetry throws distinct extract errors", async () => {
  await assert.rejects(
    parseStudyGuideJsonWithRepairAndRetry<{ a: number }>({
      raw: "no json here",
      label: "A10",
      retry: async () => "{\"a\":1}",
      maxTruncationRetries: 0,
    }),
    (err: unknown) => err instanceof Error && err.message.includes("STEP_A1_JSON_EXTRACT_FAILED"),
  );
});

test("parseStudyGuideJsonWithRepair fixes invalid escapes in strings", async () => {
  let repairCalls = 0;
  const raw = "{\"a\":\"\\q\"}";
  const parsed = await parseStudyGuideJsonWithRepair<{ a: string }>(raw, "A1", async () => {
    repairCalls += 1;
    return "{\"a\":\"ok\"}";
  });

  assert.strictEqual(repairCalls, 0);
  assert.strictEqual(parsed.a, "\\q");
});

test("parseStudyGuideJsonWithRepair handles dangling backslash in string", async () => {
  let repairCalls = 0;
  const raw = "{\"a\":\"C:\\\"}";
  const parsed = await parseStudyGuideJsonWithRepair<{ a: string }>(raw, "A2", async () => {
    repairCalls += 1;
    return "{\"a\":\"ok\"}";
  });

  assert.strictEqual(repairCalls, 0);
  assert.strictEqual(parsed.a, "C:\\");
});

test("parseStudyGuideJsonWithRepair preserves valid escapes", async () => {
  let repairCalls = 0;
  const raw = "{\"a\":\"line\\nbreak\",\"b\":\"\\u1234\"}";
  const parsed = await parseStudyGuideJsonWithRepair<{ a: string; b: string }>(raw, "A3", async () => {
    repairCalls += 1;
    return "{\"a\":\"ok\",\"b\":\"ok\"}";
  });

  assert.strictEqual(repairCalls, 0);
  assert.strictEqual(parsed.a, "line\nbreak");
  assert.strictEqual(parsed.b.length, 1);
  assert.strictEqual(parsed.b.charCodeAt(0), 0x1234);
});

test("callStudyGuideResponses JSON mode yields clean JSON", async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (input, init) => {
    const request = input instanceof Request ? input : null;
    const url = typeof input === "string" ? input : request ? request.url : "";
    if (!url.endsWith("/responses")) {
      throw new Error(`Unexpected fetch: ${url}`);
    }
    const rawBody = request ? await request.text() : typeof init?.body === "string" ? init.body : "";
    const payload = rawBody ? JSON.parse(rawBody) : {};
    assert.strictEqual(payload?.text?.format?.type, "json_object");
    assert.strictEqual("response" in payload, false);
    return new Response(JSON.stringify({ output_text: ["{\"ok\":true}"] }), {
      status: 200,
      headers: { "content-type": "application/json" },
    });
  };

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
    } as Env;
    const raw = await callStudyGuideResponses(
      env,
      "req-1",
      "B2-test",
      STUDY_GUIDE_MODEL,
      "prompt",
      200,
      { expectsJson: true },
    );
    let repairCalls = 0;
    const parsed = await parseStudyGuideJsonWithRepair<{ ok: boolean }>(raw, "B2-test", async () => {
      repairCalls += 1;
      return "{\"ok\":false}";
    });

    assert.strictEqual(raw.trim(), "{\"ok\":true}");
    assert.strictEqual(repairCalls, 0);
    assert.strictEqual(parsed.ok, true);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
