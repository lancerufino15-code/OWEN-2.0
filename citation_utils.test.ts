/// <reference types="node" />
/**
 * Tests client-side citation utilities for rendering and resolving sources.
 *
 * Used by: `npm test` to validate citation pill DOM behavior.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Placeholder replacement, pill rendering, and URL validation/disable logic.
 *
 * Assumptions:
 * - Uses JSDOM to supply DOM globals used by citation utilities.
 */

import assert from "node:assert";
import test from "node:test";
import { JSDOM } from "jsdom";
import {
  buildCitationMapFromSegments,
  buildCitationPill,
  replaceCitationPlaceholders,
  resolveCitationLinks,
} from "./public/citation_utils.mjs";

/**
 * Execute a callback with JSDOM globals configured.
 *
 * @param html - HTML string to load into the JSDOM document.
 * @param run - Callback that receives the document.
 */
function withDom(html: string, run: (doc: Document) => void) {
  const dom = new JSDOM(html);
  const globalAny = globalThis as any;
  const originalDocument = globalAny.document;
  const originalNodeFilter = globalAny.NodeFilter;
  globalAny.document = dom.window.document;
  globalAny.NodeFilter = dom.window.NodeFilter;
  try {
    run(dom.window.document);
  } finally {
    if (originalDocument) {
      globalAny.document = originalDocument;
    } else {
      delete globalAny.document;
    }
    if (originalNodeFilter) {
      globalAny.NodeFilter = originalNodeFilter;
    } else {
      delete globalAny.NodeFilter;
    }
  }
}

test("replaceCitationPlaceholders swaps placeholders with citation pills", () => {
  withDom('<div id="root">Fact __OWEN_CITE__1__ done.</div>', (document) => {
    const root = document.getElementById("root");
    assert.ok(root);
    const segments = [
      { type: "text", text: "Fact " },
      { type: "citation", id: 1, url: "https://source.test", title: "Source" },
      { type: "text", text: " done." },
    ];
    const citationMap = buildCitationMapFromSegments(segments);
    replaceCitationPlaceholders(root, citationMap, "__OWEN_CITE__");
    const link = root.querySelector("a.citation-chip");
    assert.ok(link);
    assert.strictEqual(link?.getAttribute("href"), "https://source.test");
    assert.strictEqual(link?.textContent, "[1]");
  });
});

test("buildCitationPill disables missing urls", () => {
  withDom("<div></div>", () => {
    const wrapper = buildCitationPill({ id: 2, url: "" });
    assert.strictEqual(wrapper.nodeName, "SUP");
    const button = wrapper.querySelector("button.citation-chip");
    assert.ok(button);
    assert.strictEqual(button?.getAttribute("href"), null);
    assert.strictEqual(button?.getAttribute("aria-disabled"), "true");
  });
});

test("resolveCitationLinks disables invalid urls", () => {
  withDom('<div id="root"><a class="cite" data-cite-id="1" href="#">[1]</a></div>', (document) => {
    const root = document.getElementById("root");
    assert.ok(root);
    const citationMap = new Map([["1", "/#"]]);
    resolveCitationLinks(root, citationMap);
    const link = root.querySelector("a.cite");
    assert.ok(link);
    assert.strictEqual(link?.getAttribute("href"), null);
    assert.ok(link?.classList.contains("cite--disabled"));
  });
});
