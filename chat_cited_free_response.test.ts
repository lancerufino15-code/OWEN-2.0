/// <reference types="node" />
/**
 * Tests citation handling in free-response mode.
 *
 * Used by: `npm test` to validate Responses API citation wiring.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Citation segmentation, web_search tool inclusion, and minimum-source warnings.
 *
 * Assumptions:
 * - Fetch is stubbed and does not reach the real OpenAI endpoint.
 */

import assert from "node:assert";
import test from "node:test";
import worker, { segmentWithCitationPills } from "./src/index";
import type { Env } from "./src/types";

/**
 * Build a mocked Responses API payload with URL citation annotations.
 *
 * @param sourceCount - Number of citations to include.
 * @param opts - Optional flags to include web_search source list.
 * @returns Response object with serialized JSON body.
 */
function buildCitedResponse(sourceCount: number, opts: { includeSearchSources?: boolean } = {}) {
  const includeSearchSources = opts.includeSearchSources !== false;
  const markers = Array.from({ length: sourceCount }, (_, i) => `[${i + 1}]`);
  const text = markers.length ? markers.map((marker, idx) => `Fact ${idx + 1} ${marker}`).join(" ") : "Unsourced answer.";
  const citations = markers.map((marker, idx) => {
    const start = text.indexOf(marker);
    return {
      type: "url_citation",
      start_index: start,
      end_index: start + marker.length,
      url: `https://source-${idx + 1}.test`,
      title: `Source ${idx + 1}`,
    };
  });
  const output = [
    {
      type: "message",
      role: "assistant",
      content: [
        {
          type: "output_text",
          text,
          annotations: citations,
        },
      ],
    },
  ];
  if (includeSearchSources) {
    output.push({
      type: "web_search_call",
      action: {
        sources: citations.map(entry => ({ url: entry.url, title: entry.title })),
      },
    });
  }
  return new Response(JSON.stringify({ output }), {
    status: 200,
    headers: { "content-type": "application/json" },
  });
}

test("segmentWithCitationPills sorts citations and de-duplicates URLs", () => {
  const text = "Alpha [1] Beta [2] Gamma [1]";
  const first = text.indexOf("[1]");
  const second = text.indexOf("[2]");
  const third = text.lastIndexOf("[1]");
  const citations = [
    { start_index: third, end_index: third + 3, url: "https://a.test", title: "A2" },
    { start_index: first, end_index: first + 3, url: "https://a.test", title: "A1" },
    { start_index: second, end_index: second + 3, url: "https://b.test", title: "B" },
  ];

  const result = segmentWithCitationPills(text, citations);
  const rendered = result.segments.map(seg => (seg.type === "text" ? seg.text : `#${seg.id}`));

  assert.deepStrictEqual(rendered, ["Alpha ", "#1", " Beta ", "#2", " Gamma ", "#1"]);
  assert.strictEqual(result.urlToId.get("https://a.test"), 1);
  assert.strictEqual(result.urlToId.get("https://b.test"), 2);
});

test("free-response chat uses web_search and returns citation segments", async () => {
  const originalFetch = globalThis.fetch;
  let capturedPayload: any = null;
  globalThis.fetch = async (input, init) => {
    const request = input instanceof Request ? input : null;
    const url = typeof input === "string" ? input : request ? request.url : "";
    if (!url.endsWith("/responses")) {
      throw new Error(`Unexpected fetch: ${url}`);
    }
    const rawBody = request ? await request.text() : typeof init?.body === "string" ? init.body : "";
    capturedPayload = rawBody ? JSON.parse(rawBody) : {};
    const totalSources = 8;
    const markers = Array.from({ length: totalSources }, (_, i) => `[${i + 1}]`);
    const text = markers.map((marker, idx) => `Fact ${idx + 1} ${marker}`).join(" ");
    const citations = markers.map((marker, idx) => {
      const start = text.indexOf(marker);
      return {
        type: "url_citation",
        start_index: start,
        end_index: start + marker.length,
        url: `https://source-${idx + 1}.test`,
        title: `Source ${idx + 1}`,
      };
    });
    return new Response(
      JSON.stringify({
        output: [
          {
            type: "message",
            role: "assistant",
            content: [
              {
                type: "output_text",
                text,
                annotations: citations,
              },
            ],
          },
          {
            type: "web_search_call",
            action: {
              sources: citations.map(entry => ({ url: entry.url, title: entry.title })),
            },
          },
        ],
      }),
      { status: 200, headers: { "content-type": "application/json" } },
    );
  };

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
    } as unknown as Env;

    const req = new Request("http://localhost/api/chat", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        messages: [{ role: "user", content: "Explain the topic." }],
      }),
    });

    const resp = await worker.fetch(req, env);
    const body = await resp.json();

    assert.strictEqual(body.ok, true);
    assert.ok(Array.isArray(capturedPayload?.tools));
    assert.ok(capturedPayload.tools.some((tool: any) => tool?.type === "web_search"));
    assert.ok(Array.isArray(capturedPayload?.include));
    assert.ok(capturedPayload.include.includes("web_search_call.action.sources"));
    assert.ok(Array.isArray(body.answerSegments));
    assert.ok(Array.isArray(body.sources));
    assert.deepStrictEqual(
      body.answerSegments.filter((seg: any) => seg.type === "citation").map((seg: any) => seg.id),
      [1, 2, 3, 4, 5, 6, 7, 8],
    );
    assert.deepStrictEqual(body.sources.map((src: any) => src.id), [1, 2, 3, 4, 5, 6, 7, 8]);
    assert.ok(Array.isArray(body.consultedSources));
    assert.ok(!Array.isArray(body.warnings) || body.warnings.length === 0);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("test_need_more_sources_does_not_block_answer_generation", async () => {
  const originalFetch = globalThis.fetch;
  let callCount = 0;
  let currentCount = 3;
  globalThis.fetch = async (input, init) => {
    callCount += 1;
    const request = input instanceof Request ? input : null;
    const url = typeof input === "string" ? input : request ? request.url : "";
    if (!url.endsWith("/responses")) {
      throw new Error(`Unexpected fetch: ${url}`);
    }
    return buildCitedResponse(currentCount);
  };

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
    } as unknown as Env;

    for (const count of [3, 7]) {
      currentCount = count;
      const req = new Request("http://localhost/api/chat", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          messages: [{ role: "user", content: "Explain the topic." }],
        }),
      });
      const resp = await worker.fetch(req, env);
      const body = await resp.json();

      assert.strictEqual(body.ok, true);
      assert.ok(callCount >= 1);
      assert.strictEqual(Array.isArray(body.sources) ? body.sources.length : 0, count);
      const warningCodes = Array.isArray(body.warnings) ? body.warnings.map((w: any) => w.code) : [];
      assert.ok(warningCodes.includes("INSUFFICIENT_SOURCES"));
      const text = (body.answerSegments || []).map((seg: any) => (seg.type === "text" ? seg.text : "")).join("");
      assert.ok(!text.startsWith("NEED_MORE_SOURCES:"), "Expected normal answer text");
    }
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("free-response strips citations when web_search sources are missing", async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (input, init) => {
    const request = input instanceof Request ? input : null;
    const url = typeof input === "string" ? input : request ? request.url : "";
    if (!url.endsWith("/responses")) {
      throw new Error(`Unexpected fetch: ${url}`);
    }
    return buildCitedResponse(1, { includeSearchSources: false });
  };

  try {
    const env = {
      OPENAI_API_KEY: "test-key",
      OPENAI_API_BASE: "https://api.openai.com/v1",
      ASSETS: { fetch: () => new Response("Not found", { status: 404 }) },
    } as unknown as Env;

    const req = new Request("http://localhost/api/chat", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        messages: [{ role: "user", content: "Explain the topic." }],
      }),
    });

    const resp = await worker.fetch(req, env);
    const body = await resp.json();

    assert.strictEqual(body.ok, true);
    assert.strictEqual(Array.isArray(body.sources) ? body.sources.length : 0, 0);
    const text = (body.answerSegments || []).map((seg: any) => (seg.type === "text" ? seg.text : "")).join("");
    assert.ok(!text.includes("[1]"), "Expected citation markers to be stripped");
    assert.ok(!text.includes("Sources unavailable for this response; please retry."));
    const warningCodes = Array.isArray(body.warnings) ? body.warnings.map((w: any) => w.code) : [];
    assert.ok(warningCodes.includes("NO_WEB_SOURCES"));
    assert.ok(warningCodes.includes("INSUFFICIENT_SOURCES"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});
