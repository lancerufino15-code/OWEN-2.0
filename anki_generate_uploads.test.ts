/// <reference types="node" />
/**
 * Tests the Anki upload pipeline handling for lecture assets.
 *
 * Used by: `npm test` to validate `/api/anki/generate` input handling.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Multipart field name variants, missing transcript handling, and debug payloads.
 *
 * Assumptions:
 * - Uses a minimal in-memory Env stub; no real R2/KV calls are made.
 */

import assert from "node:assert";
import test from "node:test";
import worker from "./src/index";
import type { Env } from "./src/types";

const AUTH_TOKEN = "test-token";

/**
 * Build a minimal Worker Env stub with auth + ingest bindings.
 *
 * @param token - Authorization token to mark as valid.
 * @returns Env-like object for test requests.
 */
function buildEnv(token: string) {
  return {
    ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
    DOCS_KV: {
      get: async (key: string) => (key === `faculty_session:${token}` ? "1" : null),
      put: async () => {},
      delete: async () => {},
    },
    OWEN_INGEST: {
      get: async () => null,
      put: async () => {},
      head: async () => null,
    },
  } as unknown as Env;
}

/**
 * Post an Anki generate request to the worker handler.
 *
 * @param form - Multipart form body for the Anki endpoint.
 * @param env - Worker Env bindings stub.
 * @returns Fetch Response from the worker.
 */
async function postAnki(form: FormData, env: Env) {
  const req = new Request("http://localhost/api/anki/generate", {
    method: "POST",
    headers: { authorization: `Bearer ${AUTH_TOKEN}` },
    body: form,
  });
  return worker.fetch(req, env);
}

test("anki generate accepts slideImages key without brackets", async () => {
  const env = buildEnv(AUTH_TOKEN);
  const form = new FormData();
  form.append("slideImages", new File([new Uint8Array([0x00])], "slide-1.txt", { type: "text/plain" }));
  form.append("transcriptTxt", new File(["hello"], "transcript.txt", { type: "text/plain" }));

  const resp = await postAnki(form, env);
  assert.strictEqual(resp.status, 415);
  const body = await resp.json();
  assert.strictEqual(body.stage, "validate_type");
  assert.strictEqual(body.debugCode, "anki_image_type");
  assert.strictEqual(body.received.slideImagesCount, 1);
  assert.strictEqual(body.received.slideImagesFirst.name, "slide-1.txt");
  assert.strictEqual(body.received.transcriptTxt.name, "transcript.txt");
  assert.ok(String(body.received.contentTypeHeader || "").includes("multipart/form-data"));
});

test("anki generate accepts slideImages[] key", async () => {
  const env = buildEnv(AUTH_TOKEN);
  const form = new FormData();
  form.append("slideImages[]", new File([new Uint8Array([0x01])], "slide-2.txt", { type: "text/plain" }));
  form.append("transcriptTxt", new File(["hello"], "transcript.txt", { type: "text/plain" }));

  const resp = await postAnki(form, env);
  assert.strictEqual(resp.status, 415);
  const body = await resp.json();
  assert.strictEqual(body.stage, "validate_type");
  assert.strictEqual(body.debugCode, "anki_image_type");
  assert.strictEqual(body.received.slideImagesCount, 1);
  assert.strictEqual(body.received.slideImagesFirst.name, "slide-2.txt");
});

test("anki generate accepts file fallback for slidesPdf", async () => {
  const env = buildEnv(AUTH_TOKEN);
  const form = new FormData();
  form.append("file", new File([new Uint8Array([0x00, 0x01, 0x02, 0x03])], "slides.pdf", { type: "application/pdf" }));
  form.append("transcriptTxt", new File(["hello"], "transcript.txt", { type: "text/plain" }));

  const resp = await postAnki(form, env);
  assert.strictEqual(resp.status, 415);
  const body = await resp.json();
  assert.strictEqual(body.stage, "validate_type");
  assert.strictEqual(body.debugCode, "anki_pdf_type");
  assert.strictEqual(body.received.slidesPdf.name, "slides.pdf");
  assert.strictEqual(body.received.slideImagesCount, 0);
});

test("anki generate returns debug payload when transcript is missing", async () => {
  const env = buildEnv(AUTH_TOKEN);
  const form = new FormData();
  form.append("slidesPdf", new File([new Uint8Array([0x00, 0x01, 0x02, 0x03])], "slides.pdf", { type: "application/pdf" }));

  const resp = await postAnki(form, env);
  assert.strictEqual(resp.status, 400);
  const body = await resp.json();
  assert.strictEqual(body.stage, "validate_type");
  assert.strictEqual(body.debugCode, "anki_missing_transcript");
  assert.strictEqual(body.received.slidesPdf.name, "slides.pdf");
  assert.strictEqual(body.received.transcriptTxt, null);
});
