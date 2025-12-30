/// <reference types="node" />
/**
 * Tests legacy `/api/chat` SSE streaming behavior.
 *
 * Used by: `npm test` to validate the legacy streaming response format.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - SSE envelope shape and payload normalization for `/chat/completions`.
 *
 * Assumptions:
 * - Fetch is stubbed to avoid calling external APIs.
 */

import assert from "node:assert";
import test from "node:test";
import worker from "./src/index";
import type { Env } from "./src/types";

test("legacy /api/chat streams SSE envelope", async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (input, init) => {
    const request = input instanceof Request ? input : null;
    const url = typeof input === "string" ? input : request ? request.url : "";
    if (url.endsWith("/chat/completions")) {
      const rawBody =
        request ? await request.text() : typeof init?.body === "string" ? init.body : "";
      const parsed = rawBody ? JSON.parse(rawBody) : {};
      if ("max_output_tokens" in parsed || "max_tokens" in parsed) {
        throw new Error("Unexpected token limit key in /chat/completions payload.");
      }
      if (!("max_completion_tokens" in parsed)) {
        throw new Error("Missing max_completion_tokens in /chat/completions payload.");
      }
      return new Response(
        JSON.stringify({
          choices: [{ message: { content: "hello from stub" }, finish_reason: "stop" }],
          usage: { completion_tokens: 5 },
        }),
        { status: 200, headers: { "content-type": "application/json" } },
      );
    }
    throw new Error(`Unexpected fetch: ${url}`);
  };

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      DEFAULT_MODEL: "gpt-4.1-mini",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
    } as unknown as Env;

    const req = new Request("http://localhost/api/chat", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ message: "hello" }),
    });

    const resp = await worker.fetch(req, env);
    const body = await resp.text();

    assert.ok(body.includes("event: message"));
    assert.ok(body.trimEnd().endsWith("event: done\ndata: [DONE]"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});
