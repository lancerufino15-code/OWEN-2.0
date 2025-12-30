/// <reference types="node" />
/**
 * Tests OpenAI file upload helper behavior.
 *
 * Used by: `npm test` to validate file upload request shaping.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - MIME type inference and multipart form assembly for uploads.
 *
 * Assumptions:
 * - Fetch is stubbed; no external network calls are made.
 */

import assert from "node:assert";
import test from "node:test";
import { guessMimeTypeFromFilename, uploadBytesToOpenAI } from "./src/index";
import type { Env } from "./src/types";

test("guessMimeTypeFromFilename handles common extensions", () => {
  assert.strictEqual(guessMimeTypeFromFilename("slides.PDF"), "application/pdf");
  assert.strictEqual(guessMimeTypeFromFilename("notes.txt"), "text/plain");
  assert.strictEqual(guessMimeTypeFromFilename("data.csv"), "text/csv");
  assert.strictEqual(guessMimeTypeFromFilename("deck.tsv"), "text/tab-separated-values");
  assert.strictEqual(guessMimeTypeFromFilename("image.PNG"), "image/png");
  assert.strictEqual(guessMimeTypeFromFilename("archive.bin"), "application/octet-stream");
});

test("uploadBytesToOpenAI uses inferred mime type", async () => {
  const originalFetch = globalThis.fetch;
  let seenType = "";
  let seenPurpose = "";

  globalThis.fetch = async (_input, init) => {
    const body = init?.body as FormData | undefined;
    assert.ok(body && typeof body.get === "function");
    const file = body.get("file") as File | null;
    const purpose = body.get("purpose");
    assert.ok(file && typeof file.type === "string");
    seenType = file.type;
    seenPurpose = typeof purpose === "string" ? purpose : "";
    return new Response(JSON.stringify({ id: "file-123" }), {
      status: 200,
      headers: { "content-type": "application/json" },
    });
  };

  try {
    const env = {
      OPENAI_API_BASE: "https://api.openai.com/v1",
      OPENAI_API_KEY: "test-key",
    } as Env;
    const fileId = await uploadBytesToOpenAI(env, new Uint8Array([1, 2, 3]), "slides.pdf", "assistants");
    assert.strictEqual(fileId, "file-123");
    assert.strictEqual(seenType, "application/pdf");
    assert.strictEqual(seenPurpose, "assistants");
  } finally {
    if (originalFetch) {
      globalThis.fetch = originalFetch;
    } else {
      delete (globalThis as any).fetch;
    }
  }
});
