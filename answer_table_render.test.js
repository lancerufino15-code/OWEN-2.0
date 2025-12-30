/**
 * Tests table parsing and rendering for answer content blocks.
 *
 * Used by: `npm test` to validate table rendering helpers in the UI layer.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Markdown/GFM tables, pipe-only tables, and inline HTML table blocks.
 *
 * Assumptions:
 * - Uses JSDOM to supply DOM globals required by table renderers.
 */
import assert from "node:assert/strict";
import { test } from "node:test";
import { JSDOM } from "jsdom";
import { renderHtmlTable, renderMarkdownTable } from "./public/table_parser.mjs";

/**
 * Run a callback with JSDOM-provided globals.
 *
 * @param {Function} run - Callback invoked with the JSDOM document.
 */
function withDom(run) {
  const dom = new JSDOM("<!doctype html><html><body></body></html>");
  const globalAny = globalThis;
  const originalDocument = globalAny.document;
  const originalWindow = globalAny.window;
  globalAny.document = dom.window.document;
  globalAny.window = dom.window;
  try {
    run(dom.window.document);
  } finally {
    if (originalDocument) {
      globalAny.document = originalDocument;
    } else {
      delete globalAny.document;
    }
    if (originalWindow) {
      globalAny.window = originalWindow;
    } else {
      delete globalAny.window;
    }
  }
}

test("renderMarkdownTable handles GFM markdown tables", () => {
  const block = [
    "Caption: Antibiotic classes and mechanisms.",
    "| Class | Mechanism | Examples |",
    "| --- | --- | --- |",
    "| Beta-lactams | Inhibit cell wall | Penicillin [1] |",
    "| Macrolides | 50S inhibitor | Azithromycin |",
  ].join("\n");
  withDom(() => {
    const rendered = renderMarkdownTable(block, {
      formatCell: text => text,
      pickTone: () => "clinical",
    });
    assert.ok(rendered?.node);
    const table = rendered.node;
    assert.equal(table.tagName, "TABLE");
    assert.equal(table.dataset.answerTable, "true");
    assert.ok(table.classList.contains("AnswerTable"));
    assert.equal(table.querySelectorAll("thead th").length, 3);
    assert.ok(table.querySelectorAll("tbody tr").length >= 2);
  });
});

test("renderMarkdownTable handles pipe-delimited tables without dividers", () => {
  const block = [
    "Caption: Rapid approach table.",
    "| Step 1 | Start empiric coverage |",
    "| Step 2 | Narrow by culture |",
    "| Step 3 | Monitor response |",
  ].join("\n");
  withDom(() => {
    const rendered = renderMarkdownTable(block, {
      formatCell: text => text,
      pickTone: () => "clinical",
    });
    assert.ok(rendered?.node);
    const table = rendered.node;
    assert.equal(table.tagName, "TABLE");
    assert.ok(table.querySelectorAll("tbody tr").length >= 2);
  });
});

test("renders markdown, html, and pipe tables from a single answer", () => {
  const markdownTable = [
    "Caption: Antibiotic classes and mechanisms.",
    "| Class | Mechanism | Examples |",
    "| --- | --- | --- |",
    "| Beta-lactams | Inhibit cell wall | Penicillin [1] |",
    "| Macrolides | 50S inhibitor | Azithromycin |",
  ].join("\n");
  const htmlTable = [
    "<table>",
    "<caption>Key labs</caption>",
    "<thead><tr><th>Lab</th><th>Value</th></tr></thead>",
    "<tbody><tr><td>Na</td><td>140</td></tr><tr><td>K</td><td>4.2</td></tr></tbody>",
    "</table>",
  ].join("");
  const pipeTable = [
    "Caption: Rapid approach table.",
    "| Step 1 | Start empiric coverage |",
    "| Step 2 | Narrow by culture |",
    "| Step 3 | Monitor response |",
  ].join("\n");
  const answer = [markdownTable, htmlTable, pipeTable].join("\n\n");

  withDom((doc) => {
    const container = doc.createElement("div");
    answer.split(/\n{2,}/).forEach(block => {
      const trimmed = (block || "").trim();
      if (!trimmed) return;
      const htmlRendered = renderHtmlTable(trimmed, {
        formatCell: text => text,
        pickTone: () => "clinical",
      });
      if (htmlRendered?.node) {
        container.appendChild(htmlRendered.node);
        return;
      }
      const markdownRendered = renderMarkdownTable(trimmed, {
        formatCell: text => text,
        pickTone: () => "clinical",
      });
      if (markdownRendered?.node) container.appendChild(markdownRendered.node);
    });

    const tables = container.querySelectorAll("table");
    assert.equal(tables.length, 3);
    tables.forEach(table => {
      assert.ok(table.querySelectorAll("tr").length >= 2);
      assert.ok(table.querySelectorAll("td, th").length >= 2);
    });
  });
});
