/// <reference types="node" />
/**
 * Tests OpenAI payload sanitization and retry logic.
 *
 * Used by: `npm test` to validate request shaping for OpenAI-compatible APIs.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Stripping unsupported params and retrying when the API rejects payloads.
 *
 * Assumptions:
 * - Retry behavior is deterministic and limited to a single fallback attempt.
 */

import assert from "node:assert";
import test from "node:test";
import { sanitizeOpenAIPayload, sendOpenAIWithUnsupportedParamRetry } from "./src/index";

test("sanitizeOpenAIPayload strips seed for all models", () => {
  const payload = {
    model: "gpt-4o",
    temperature: 0,
    top_p: 1,
    seed: 42,
    response: { format: "json_object" },
  };
  const result = sanitizeOpenAIPayload(payload, { endpoint: "responses" });
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "seed"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "response"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "temperature"), true);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "top_p"), true);
});

test("sanitizeOpenAIPayload drops sampling params for unsupported models", () => {
  const payload = {
    model: "gpt-5-mini",
    temperature: 0,
    top_p: 1,
    seed: 42,
    response: { format: "json_object" },
  };
  const result = sanitizeOpenAIPayload(payload, { endpoint: "responses" });
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "seed"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "response"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "temperature"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "top_p"), false);
});

test("sanitizeOpenAIPayload honors OWEN_STRIP_SAMPLING_PARAMS", () => {
  const payload = {
    model: "gpt-4o",
    temperature: 0,
    top_p: 1,
    seed: 42,
    response: { format: "json_object" },
  };
  const result = sanitizeOpenAIPayload(payload, {
    endpoint: "responses",
    env: { OWEN_STRIP_SAMPLING_PARAMS: "1" },
  });
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "seed"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "response"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "temperature"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(result.payload, "top_p"), false);
});

test("sendOpenAIWithUnsupportedParamRetry retries once without temperature/top_p", async () => {
  const calls: Array<Record<string, unknown>> = [];
  const result = await sendOpenAIWithUnsupportedParamRetry({
    payload: {
      model: "gpt-4o",
      temperature: 0,
      top_p: 1,
      seed: 42,
      response: { format: "json_object" },
    },
    endpoint: "responses",
    label: "test",
    send: async (payload) => {
      calls.push(payload);
      if (calls.length === 1) {
        return { ok: false, errorText: "Unsupported parameter: 'temperature'." };
      }
      return { ok: true, value: "ok" };
    },
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.attempts, 2);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(calls[0], "seed"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(calls[0], "response"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(calls[0], "temperature"), true);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(calls[0], "top_p"), true);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(calls[1], "temperature"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(calls[1], "top_p"), false);
});

test("sendOpenAIWithUnsupportedParamRetry retries once on unknown response parameter", async () => {
  const calls: Array<Record<string, unknown>> = [];
  const result = await sendOpenAIWithUnsupportedParamRetry({
    payload: {
      model: "gpt-4o",
      response: { format: "json_object" },
    },
    endpoint: "responses",
    label: "test-response",
    send: async (payload) => {
      calls.push(payload);
      if (calls.length === 1) {
        return { ok: false, errorText: "Unknown parameter: 'response'." };
      }
      return { ok: true, value: "ok" };
    },
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.attempts, 2);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(calls[0], "response"), false);
  assert.strictEqual(Object.prototype.hasOwnProperty.call(calls[1], "response"), false);
});
