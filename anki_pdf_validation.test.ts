/// <reference types="node" />
/**
 * Tests PDF signature validation for Anki uploads.
 *
 * Used by: `npm test` to validate `isAnkiPdfFile` behavior.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - PDF magic-byte detection and rejection of empty files.
 *
 * Assumptions:
 * - Uses the File API provided by the test runtime (node:test + undici).
 */

import assert from "node:assert";
import test from "node:test";
import { isAnkiPdfFile } from "./src/index";

test("isAnkiPdfFile validates PDF signature bytes", async () => {
  const good = new File(
    [new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x37])],
    "slides.pdf",
    { type: "application/pdf" },
  );
  const bad = new File([new Uint8Array([0x00, 0x01, 0x02, 0x03, 0x04])], "slides.pdf", {
    type: "application/pdf",
  });
  assert.strictEqual(await isAnkiPdfFile(good), true);
  assert.strictEqual(await isAnkiPdfFile(bad), false);
});

test("isAnkiPdfFile rejects empty files", async () => {
  const empty = new File([new Uint8Array()], "slides.pdf", { type: "application/pdf" });
  assert.strictEqual(await isAnkiPdfFile(empty), false);
});
