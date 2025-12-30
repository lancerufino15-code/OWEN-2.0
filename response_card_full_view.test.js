/**
 * Tests full response card rendering and DOM behavior in JSDOM.
 *
 * Used by: `npm test` to validate response card UI flows.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Full-view modal rendering, DOM wiring, and UI defaults.
 *
 * Assumptions:
 * - Runs in JSDOM with `public/index.html` as the DOM fixture.
 */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";
import { JSDOM } from "jsdom";

const html = readFileSync(new URL("./public/index.html", import.meta.url), "utf8");

/**
 * Execute a callback with JSDOM globals installed.
 *
 * @param {Function} run - Async callback invoked with the JSDOM window.
 */
async function withDom(run) {
  const dom = new JSDOM(html, { url: "http://localhost" });
  const globalAny = globalThis;
  const originalDocument = globalAny.document;
  const originalWindow = globalAny.window;
  const originalNavigator = globalAny.navigator;
  const originalHTMLElement = globalAny.HTMLElement;
  const originalNode = globalAny.Node;
  const originalFetch = globalAny.fetch;

  globalAny.document = dom.window.document;
  globalAny.window = dom.window;
  globalAny.navigator = dom.window.navigator;
  globalAny.HTMLElement = dom.window.HTMLElement;
  globalAny.Node = dom.window.Node;
  globalAny.fetch = async (url) => {
    const href = typeof url === "string" ? url : url?.url || "";
    const payload = href.includes("/api/library/list")
      ? { items: [] }
      : href.includes("/api/library/search")
        ? { results: [] }
        : {};
    return {
      ok: true,
      json: async () => payload,
      text: async () => "",
    };
  };

  try {
    await run(dom.window);
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
    if (originalNavigator) {
      globalAny.navigator = originalNavigator;
    } else {
      delete globalAny.navigator;
    }
    if (originalHTMLElement) {
      globalAny.HTMLElement = originalHTMLElement;
    } else {
      delete globalAny.HTMLElement;
    }
    if (originalNode) {
      globalAny.Node = originalNode;
    } else {
      delete globalAny.Node;
    }
    if (originalFetch) {
      globalAny.fetch = originalFetch;
    } else {
      delete globalAny.fetch;
    }
  }
}

test("answer renders fully by default without view toggles or accordions", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = [
      "Response:",
      "| Antibiotic | Class | Notes |",
      "| --- | --- | --- |",
      "| Nitrofurantoin | Nitrofurans | First-line for uncomplicated UTI |",
      "| Trimethoprim | Folate antagonist | Alternative option |",
      "",
      "Plan:",
      "- Monitor renal function",
      "- Avoid in late pregnancy",
      "",
      "Warnings:",
      "- Consider local resistance patterns",
      "- Adjust dose for renal impairment",
    ].join("\n");

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const responseContent = card.querySelector(".response-content");
    assert.ok(responseContent);
    assert.ok(!responseContent.classList.contains("truncate"));
    assert.ok(!Array.from(responseContent.classList).some(cls => cls.startsWith("line-clamp")));
    assert.equal(card.querySelectorAll("details.response-section").length, 0);
    assert.equal(card.querySelectorAll(".response-view-toggle").length, 0);
    assert.equal(card.querySelectorAll(".toc-action").length, 0);
    assert.equal(card.querySelectorAll(".response-summary").length, 0);
    assert.ok(card.querySelector(".response-collapse-toggle"));

    const body = card.querySelector(".response-body");
    assert.ok(body);
    assert.equal(card.dataset.collapsed, "false");
    assert.equal(body.getAttribute("aria-hidden"), "false");

    const toolbarLabels = Array.from(card.querySelectorAll(".response-toolbar button"))
      .map(btn => (btn.textContent || "").trim());
    assert.ok(toolbarLabels.includes("Summary"));
    assert.ok(!toolbarLabels.includes("Full"));
    assert.ok(!toolbarLabels.includes("Expand all"));
    assert.ok(!toolbarLabels.includes("Collapse all"));

    const table = card.querySelector("table");
    assert.ok(table);

    const sectionBodyText = card.querySelector(".section-body")?.textContent || "";
    assert.ok(sectionBodyText.includes("Trimethoprim"));
  });
});

test("summary toggle keeps stored answer text unchanged", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard, getResponseState } = await import("./public/chat.js");
    const answer = [
      "## Overview",
      "Core takeaways live here. They should remain unchanged.",
      "",
      "- Bullet one",
      "- Bullet two",
    ].join("\n");

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const summaryBtn = card.querySelector("[data-action=\"toggle-summary\"]");
    assert.ok(summaryBtn);
    const collapse = card.querySelector(".response-collapse-toggle");
    assert.ok(collapse);

    assert.equal(getResponseState(card)?.answerText, answer);
    summaryBtn.click();
    assert.equal(card.dataset.viewMode, "summary");
    assert.ok(card.querySelector(".response-summary"));
    assert.equal(getResponseState(card)?.answerText, answer);

    collapse.click();
    assert.equal(card.dataset.collapsed, "true");
    assert.equal(getResponseState(card)?.answerText, answer);

    collapse.click();
    summaryBtn.click();
    assert.equal(card.dataset.viewMode, "full");
    assert.equal(getResponseState(card)?.answerText, answer);
  });
});

test("response header chevron collapses and expands the full answer body", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = ["## A", "Hello", "", "## B", "World"].join("\n");

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const toggle = card.querySelector(".response-collapse-toggle");
    assert.ok(toggle);

    const body = card.querySelector(".response-body");
    assert.ok(body);
    assert.equal(toggle.getAttribute("aria-expanded"), "true");
    assert.equal(card.dataset.collapsed, "false");
    assert.equal(body.getAttribute("aria-hidden"), "false");

    toggle.click();
    assert.equal(toggle.getAttribute("aria-expanded"), "false");
    assert.equal(card.dataset.collapsed, "true");
    assert.equal(body.getAttribute("aria-hidden"), "true");

    toggle.click();
    assert.equal(toggle.getAttribute("aria-expanded"), "true");
    assert.equal(card.dataset.collapsed, "false");
    assert.equal(body.getAttribute("aria-hidden"), "false");
  });
});

test("answer renders even when sources are missing", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = "Fact [1] should still render even without sources.";

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer, sources: [] });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const content = card.querySelector(".response-content");
    assert.ok(content);
    const text = (content.textContent || "").trim();
    assert.ok(text.includes("Fact"));
    assert.ok(!text.includes("Sources unavailable for this response"));

    const sourcesLabel = card.querySelector("[data-role=\"sources-count\"]");
    assert.ok(sourcesLabel);
    assert.equal((sourcesLabel.textContent || "").trim(), "Sources: 0");

    const viewBtn = card.querySelector("[data-action=\"view-sources\"]");
    assert.ok(viewBtn);
    assert.equal(viewBtn.disabled, true);
    assert.equal(viewBtn.title, "No sources returned for this response.");
  });
});

test("inline references enable drawer with unverified label", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = [
      "Key points go here.",
      "",
      "References:",
      "1. Alpha Study",
      "2) Beta Report",
      "3 - Gamma Paper",
      "",
      "Sources unavailable for this response; please retry.",
    ].join("\n");

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer, sources: [] });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const content = card.querySelector(".response-content");
    assert.ok(content);
    assert.ok(!(content.textContent || "").includes("Sources unavailable for this response"));

    const sourcesLabel = card.querySelector("[data-role=\"sources-count\"]");
    assert.ok(sourcesLabel);
    assert.equal((sourcesLabel.textContent || "").trim(), "References: 3");

    const tag = card.querySelector("[data-role=\"sources-tag\"]");
    assert.ok(tag);
    assert.equal(tag.hidden, false);
    assert.equal((tag.textContent || "").trim(), "Unverified");

    const viewBtn = card.querySelector("[data-action=\"view-sources\"]");
    assert.ok(viewBtn);
    assert.equal(viewBtn.disabled, false);

    viewBtn.click();
    const drawer = document.getElementById("sourcesDrawer");
    assert.ok(drawer);
    assert.ok(drawer.classList.contains("is-open"));
    const list = document.getElementById("sourcesList");
    assert.ok(list);
    const drawerText = list.textContent || "";
    assert.ok(drawerText.includes("References from answer (unverified)"));
    assert.ok(drawerText.includes("Alpha Study"));
    assert.ok(drawerText.includes("Beta Report"));
  });
});

test("no sources or references disables view sources", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = ["Just an answer.", "", "Sources unavailable for this response; please retry."].join("\n");

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer, sources: [] });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const content = card.querySelector(".response-content");
    assert.ok(content);
    assert.ok(!(content.textContent || "").includes("Sources unavailable for this response"));

    const sourcesLabel = card.querySelector("[data-role=\"sources-count\"]");
    assert.ok(sourcesLabel);
    assert.equal((sourcesLabel.textContent || "").trim(), "Sources: 0");

    const viewBtn = card.querySelector("[data-action=\"view-sources\"]");
    assert.ok(viewBtn);
    assert.equal(viewBtn.disabled, true);
    assert.equal(viewBtn.title, "No sources returned for this response.");
  });
});

test("verified sources keep existing header and drawer behavior", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = "Verified fact [1].";
    const sources = [{ id: 1, url: "https://openai.com", title: "OpenAI" }];

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer, sources });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const sourcesLabel = card.querySelector("[data-role=\"sources-count\"]");
    assert.ok(sourcesLabel);
    assert.equal((sourcesLabel.textContent || "").trim(), "Sources: 1");

    const tag = card.querySelector("[data-role=\"sources-tag\"]");
    assert.ok(tag);
    assert.equal(tag.hidden, true);

    const viewBtn = card.querySelector("[data-action=\"view-sources\"]");
    assert.ok(viewBtn);
    assert.equal(viewBtn.disabled, false);

    viewBtn.click();
    const list = document.getElementById("sourcesList");
    assert.ok(list);
    const drawerText = list.textContent || "";
    assert.ok(drawerText.includes("OpenAI"));
    assert.ok(!drawerText.includes("References from answer"));
  });
});

test("full response view renders markdown headings and toc chips", async () => {
  await withDom(async (window) => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = [
      "## Overview",
      "Core takeaways live here.",
      "",
      "## Dosing",
      "- Start low",
      "- Titrate slowly",
      "",
      "### Pediatric",
      "Adjust for weight.",
      "",
      "## Monitoring",
      "Check labs weekly.",
    ].join("\n");

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const sectionHeadings = Array.from(card.querySelectorAll(".section-heading"));
    assert.ok(sectionHeadings.length >= 3);
    sectionHeadings.forEach(heading => {
      assert.ok(["H2", "H3"].includes(heading.tagName));
      assert.ok(heading.id);
    });

    const tocItems = Array.from(card.querySelectorAll(".toc-item"));
    assert.ok(tocItems.length >= 2);

    let scrolledTo = null;
    window.HTMLElement.prototype.scrollIntoView = function () {
      scrolledTo = this;
    };
    tocItems[0].click();
    assert.ok(scrolledTo);
    assert.equal(scrolledTo.id, tocItems[0].dataset.targetId);
    assert.equal(window.location.hash, `#${tocItems[0].dataset.targetId}`);
  });
});

test("tables render heading captions and keep trailing content", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = [
      "## Warfarin",
      "| Drug | Antidote |",
      "| --- | --- |",
      "| Warfarin | Vitamin K |",
    ].join("\n");

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const heading = card.querySelector("h2.section-heading, h3.section-heading");
    assert.ok(heading);
    assert.ok((heading.textContent || "").includes("Warfarin"));
    const table = card.querySelector("table");
    assert.ok(table);
    assert.ok(heading.compareDocumentPosition(table) & Node.DOCUMENT_POSITION_FOLLOWING);
  });
});

test("section ids are slugified and deduped", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard } = await import("./public/chat.js");
    const answer = [
      "## 1. Warfarin (Vitamin K Antagonist)",
      "First section body.",
      "",
      "## 1. Warfarin (Vitamin K Antagonist)",
      "Second section body.",
    ].join("\n");

    const card = decorateAssistantResponse(answer, new Map(), { answerText: answer, msgId: "test" });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const ids = Array.from(card.querySelectorAll(".response-section .section-heading"))
      .map(node => node.id)
      .filter(Boolean);
    assert.ok(ids.length >= 2);
    assert.ok(ids[0].includes("1-warfarin-vitamin-k-antagonist"));
    assert.notEqual(ids[0], ids[1]);
    assert.ok(ids[1].endsWith("-2"));
    assert.ok(!/[()]/.test(ids[0]));
  });
});

test("stored citations hydrate inline links", async () => {
  await withDom(async () => {
    const { decorateAssistantResponse, finalizeResponseCard, buildCitationMapFromStored } = await import("./public/chat.js");
    const answer = "Evidence [1] should link to the stored source.";
    const storedSources = [{ id: 1, url: "https://source-1.test", title: "Example Source" }];
    const citationMap = buildCitationMapFromStored({ sources: storedSources });

    const card = decorateAssistantResponse(answer, citationMap, {
      answerText: answer,
      sources: storedSources,
      preserveCitationMap: true,
    });
    document.body.appendChild(card);
    finalizeResponseCard(card);

    const chip = card.querySelector(".citation-chip[data-cite-id=\"1\"]");
    assert.ok(chip);
    assert.equal(chip.getAttribute("href"), "https://source-1.test");
  });
});
