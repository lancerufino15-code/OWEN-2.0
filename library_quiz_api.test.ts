/// <reference types="node" />
/**
 * Tests `/api/library/quiz` API behavior with schema-constrained AI output.
 *
 * Used by: `npm test` to validate quiz error handling and AI bindings.
 */
import assert from "node:assert/strict";
import test from "node:test";
import worker from "./src/index";
import type { Env } from "./src/types";
import { buildExtractedKeyForHash } from "./src/pdf/cache";

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

function buildEnv(opts: {
  bucket: MemoryR2Bucket;
  aiRun?: (model: string, input: any) => Promise<any>;
  timeoutMs?: string;
  includeAi?: boolean;
}) {
  const includeAi = opts.includeAi !== false;
  const aiRun = opts.aiRun || (async () => ({ response: {} }));
  return {
    ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
    OWEN_INGEST: opts.bucket,
    DOCS_KV: { get: async () => null, put: async () => {}, delete: async () => {} },
    AI: includeAi ? { run: aiRun } : undefined,
    QUIZ_AI_TIMEOUT_MS: opts.timeoutMs,
  } as unknown as Env;
}

async function seedLecture(bucket: MemoryR2Bucket, docId: string, text: string) {
  const extractedKey = buildExtractedKeyForHash(docId);
  await bucket.put(extractedKey, text, { httpMetadata: { contentType: "text/plain; charset=utf-8" } });
}

function buildQuizRequest(docId = "doc-1") {
  return {
    docId,
    lectureTitle: "Lecture One",
  };
}

function buildQuizBatch(docId = "doc-1") {
  const makeQuestion = (id: string, answer: string) => ({
    id,
    stem: `Question ${id}?`,
    choices: [
      { id: "A", text: "Option A" },
      { id: "B", text: "Option B" },
      { id: "C", text: "Option C" },
      { id: "D", text: "Option D" },
    ],
    answer,
    rationale: "Because.",
    tags: ["tag"],
    difficulty: "easy",
    references: ["Slide 1"],
  });
  return {
    lectureId: docId,
    lectureTitle: "Lecture One",
    setSize: 5,
    questions: [
      makeQuestion("q1", "A"),
      makeQuestion("q2", "B"),
      makeQuestion("q3", "C"),
      makeQuestion("q4", "D"),
      makeQuestion("q5", "A"),
    ],
  };
}

test("library quiz rejects invalid JSON body", async () => {
  const bucket = new MemoryR2Bucket();
  let aiCalled = false;
  const env = buildEnv({
    bucket,
    aiRun: async () => {
      aiCalled = true;
      return { response: {} };
    },
  });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: "{bad json",
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 400);
  assert.strictEqual(body.error, "bad_request_invalid_json");
  assert.strictEqual(aiCalled, false);
});

test("library quiz test env includes AI binding with aiRun", async () => {
  const bucket = new MemoryR2Bucket();
  const env = buildEnv({
    bucket,
    aiRun: async () => ({ response: {} }),
  });
  assert.ok(env.AI);
  assert.strictEqual(typeof env.AI?.run, "function");
});

test("library quiz returns 502 when AI binding is missing", async () => {
  const bucket = new MemoryR2Bucket();
  await seedLecture(bucket, "doc-1", "A".repeat(400));
  const env = buildEnv({ bucket, includeAi: false });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(buildQuizRequest()),
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 502);
  assert.strictEqual(body.error, "upstream_ai_unavailable");
  assert.strictEqual(body.message, "AI binding not configured.");
});

test("library quiz returns 422 when AI returns prose", async () => {
  const bucket = new MemoryR2Bucket();
  await seedLecture(bucket, "doc-1", "A".repeat(400));
  let seenPayload: any = null;
  const env = buildEnv({
    bucket,
    aiRun: async (_model, input) => {
      seenPayload = input;
      return { response: "Here is a quiz:" };
    },
  });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(buildQuizRequest()),
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 422);
  assert.strictEqual(body.error, "quiz_invalid_json");
  assert.strictEqual(body.details, "no-json");
  assert.strictEqual(seenPayload?.response_format?.type, "json_schema");
  assert.ok(!("stream" in (seenPayload || {})));
});

test("library quiz returns 200 for valid schema output", async () => {
  const bucket = new MemoryR2Bucket();
  await seedLecture(bucket, "doc-1", "B".repeat(400));
  const expected = buildQuizBatch("doc-1");
  const env = buildEnv({
    bucket,
    aiRun: async () => ({ response: expected }),
  });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(buildQuizRequest("doc-1")),
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 200);
  assert.strictEqual(body.ok, true);
  assert.deepStrictEqual(body.batch, expected);
});

test("library quiz retries with json_object when json mode couldn't be met", async () => {
  const bucket = new MemoryR2Bucket();
  await seedLecture(bucket, "doc-1", "B".repeat(400));
  const expected = buildQuizBatch("doc-1");
  const payloads: any[] = [];
  let calls = 0;
  const env = buildEnv({
    bucket,
    aiRun: async (_model, input) => {
      payloads.push(input);
      calls += 1;
      if (calls === 1) {
        throw new Error("JSON Mode couldn't be met");
      }
      return { response: expected };
    },
  });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(buildQuizRequest("doc-1")),
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 200);
  assert.strictEqual(body.ok, true);
  assert.deepStrictEqual(body.batch, expected);
  assert.strictEqual(payloads.length, 2);
  assert.strictEqual(payloads[0]?.response_format?.type, "json_schema");
  assert.strictEqual(payloads[1]?.response_format?.type, "json_object");
  assert.strictEqual(payloads[1]?.temperature, 0.0);
});

test("library quiz accepts quiz JSON embedded in choices output", async () => {
  const bucket = new MemoryR2Bucket();
  await seedLecture(bucket, "doc-1", "B".repeat(400));
  const expected = buildQuizBatch("doc-1");
  const env = buildEnv({
    bucket,
    aiRun: async () => ({
      response: {
        choices: [
          {
            message: { content: JSON.stringify(expected) },
          },
        ],
      },
    }),
  });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(buildQuizRequest("doc-1")),
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 200);
  assert.strictEqual(body.ok, true);
  assert.deepStrictEqual(body.batch, expected);
});

test("library quiz returns 422 when schema validation fails", async () => {
  const bucket = new MemoryR2Bucket();
  await seedLecture(bucket, "doc-1", "C".repeat(400));
  const env = buildEnv({
    bucket,
    aiRun: async () => ({ response: { lectureTitle: "Lecture One", setSize: 5, questions: [] } }),
  });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(buildQuizRequest("doc-1")),
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 422);
  assert.strictEqual(body.error, "quiz_invalid_json");
  assert.strictEqual(body.details, "schema");
  assert.ok(Array.isArray(body.validationErrors));
});

test("library quiz returns 502 when AI throws non-timeout error", async () => {
  const bucket = new MemoryR2Bucket();
  await seedLecture(bucket, "doc-1", "C".repeat(400));
  const env = buildEnv({
    bucket,
    aiRun: async () => {
      throw new Error("AI blew up");
    },
  });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(buildQuizRequest("doc-1")),
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 502);
  assert.strictEqual(body.error, "upstream_ai_error");
  assert.strictEqual(body.message, "AI blew up");
});

test("library quiz returns 504 on AI timeout", async () => {
  const bucket = new MemoryR2Bucket();
  await seedLecture(bucket, "doc-1", "D".repeat(400));
  const env = buildEnv({
    bucket,
    timeoutMs: "5",
    aiRun: async () => new Promise(() => {}),
  });

  const req = new Request("http://localhost/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(buildQuizRequest("doc-1")),
  });
  const resp = await worker.fetch(req, env);
  const body = await resp.json();

  assert.strictEqual(resp.status, 504);
  assert.strictEqual(body.error, "upstream_timeout");
});
