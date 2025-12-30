/**
 * Tests quiz window controls and DOM interactions in the chat UI.
 *
 * Used by: `npm test` to validate quiz modal behavior and state transitions.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Opening/closing, interrupt handling, and multi-set quiz flows.
 *
 * Assumptions:
 * - Runs in JSDOM with `public/index.html` as the UI fixture.
 */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";
import { JSDOM } from "jsdom";

const html = readFileSync(new URL("./public/index.html", import.meta.url), "utf8");

/**
 * Build a fetch-like JSON response stub.
 *
 * @param {any} payload - JSON payload to return.
 * @param {boolean} ok - Response ok flag.
 * @param {number} status - HTTP status code.
 * @returns {object} Response-like stub with json/text helpers.
 */
function makeJsonResponse(payload, ok = true, status = 200) {
  return {
    ok,
    status,
    json: async () => payload,
    text: async () => JSON.stringify(payload),
  };
}

/**
 * Build a delayed JSON response that can be resolved manually.
 *
 * @param {any} payload - JSON payload to return on resolve.
 * @returns {object} Response-like stub with a resolve hook.
 */
function makeDelayedJsonResponse(payload) {
  let resolver = null;
  let resolved = false;
  return {
    ok: true,
    status: 200,
    json: async () => payload,
    text: async () => new Promise(resolve => {
      resolver = resolve;
      if (resolved) {
        resolve(JSON.stringify(payload));
      }
    }),
    resolve: () => {
      if (resolver) {
        resolver(JSON.stringify(payload));
      } else {
        resolved = true;
      }
    },
  };
}

/**
 * Poll until a predicate is true or timeout occurs.
 *
 * @param {Function} predicate - Condition to satisfy.
 * @param {number} ticks - Maximum async ticks to wait.
 * @returns {Promise<void>} Resolves when predicate is true.
 * @throws If the predicate never returns true within the tick limit.
 */
async function waitFor(predicate, ticks = 20) {
  for (let i = 0; i < ticks; i += 1) {
    if (predicate()) return;
    await new Promise(resolve => setTimeout(resolve, 0));
  }
  throw new Error("Timed out waiting for condition.");
}

/**
 * Build a quiz batch fixture with 5 questions.
 *
 * @param {string} lectureTitle - Title used in the batch.
 * @param {number} startIndex - Starting index for question ids.
 * @returns {object} Quiz batch payload.
 */
function buildQuizBatch(lectureTitle, startIndex = 1) {
  const questions = Array.from({ length: 5 }, (_, idx) => {
    const idNumber = startIndex + idx;
    return {
      id: `q${idNumber}`,
      stem: `Question ${idNumber}?`,
      choices: [
        { id: "A", text: "Option A" },
        { id: "B", text: "Option B" },
        { id: "C", text: "Option C" },
        { id: "D", text: "Option D" },
      ],
      answer: "A",
      rationale: `Reason ${idNumber}.`,
      tags: ["tag"],
      difficulty: "easy",
      references: [`Slide ${idNumber}`],
    };
  });
  return {
    lectureTitle,
    setSize: 5,
    questions,
  };
}

/**
 * Click through all correct answers for a given quiz set in the DOM.
 *
 * @param {Window} win - JSDOM window instance.
 * @param {number} setIndex - Quiz set index to complete.
 */
function completeQuizSet(win, setIndex = 0) {
  const selector = `.quiz-option[data-quiz-set-index="${setIndex}"][data-choice-id="A"]`;
  Array.from(win.document.querySelectorAll(selector)).forEach(btn => btn.click());
}

test("lecture quiz window opens on demand, closes, and handles interrupt outcomes", async () => {
  const dom = new JSDOM(html, { url: "http://localhost/student" });
  const globalAny = globalThis;
  const originalDocument = globalAny.document;
  const originalWindow = globalAny.window;
  const originalNavigator = globalAny.navigator;
  const originalHTMLElement = globalAny.HTMLElement;
  const originalNode = globalAny.Node;
  const originalFetch = globalAny.fetch;
  const originalAlert = globalAny.alert;

  const fetchCalls = [];
  let quizCallCount = 0;
  let interruptShouldFail = false;
  let delayedQuizSuccess = null;
  let delayedQuizFailure = null;

  const quizBatch1 = buildQuizBatch("Lecture One", 1);
  const quizBatch2 = buildQuizBatch("Lecture One", 6);

  globalAny.document = dom.window.document;
  globalAny.window = dom.window;
  globalAny.navigator = dom.window.navigator;
  globalAny.HTMLElement = dom.window.HTMLElement;
  globalAny.Node = dom.window.Node;
  const alertStub = () => {};
  globalAny.alert = alertStub;
  dom.window.alert = alertStub;
  dom.window.matchMedia = () => ({
    matches: false,
    addListener: () => {},
    removeListener: () => {},
    addEventListener: () => {},
    removeEventListener: () => {},
  });
  globalAny.fetch = async (input) => {
    const url = typeof input === "string" ? input : input?.url || "";
    fetchCalls.push(url);
    if (url.startsWith("/api/library/list")) {
      return makeJsonResponse({
        items: [
          {
            docId: "doc-1",
            title: "Lecture One",
            status: "missing",
            bucket: "owen-ingest",
            key: "library/lecture-one.pdf",
          },
        ],
      });
    }
    if (url.startsWith("/api/library/search")) {
      return makeJsonResponse({ results: [] });
    }
    if (url.startsWith("/api/library/ingest")) {
      return makeJsonResponse({ status: "ready", docId: "doc-1", extractedKey: "extracted" });
    }
    if (url.startsWith("/api/library/quiz/interrupt")) {
      if (interruptShouldFail) {
        return makeJsonResponse({ ok: false, message: "Quiz interrupt failed." }, false, 500);
      }
      return makeJsonResponse({ ok: true });
    }
    if (url.startsWith("/api/library/quiz")) {
      quizCallCount += 1;
      if (quizCallCount === 1) {
        return makeJsonResponse({ ok: true, batch: quizBatch1 });
      }
      if (quizCallCount === 2) {
        delayedQuizSuccess = makeDelayedJsonResponse({ ok: true, batch: quizBatch2 });
        return delayedQuizSuccess;
      }
      if (quizCallCount === 3) {
        return makeJsonResponse({ ok: true, batch: quizBatch1 });
      }
      if (quizCallCount === 4) {
        delayedQuizFailure = makeDelayedJsonResponse({ ok: true, batch: quizBatch2 });
        return delayedQuizFailure;
      }
      return makeJsonResponse({ ok: true, batch: quizBatch2 });
    }
    if (url.startsWith("/api/conversations")) {
      return makeJsonResponse({ conversations: [] });
    }
    return makeJsonResponse({});
  };

  try {
    await import("./public/chat.js");
    await waitFor(() => dom.window.document.querySelector("#librarySelect option[value=\"doc-1\"]"));

    const quizModal = dom.window.document.querySelector("#quizModal");
    const quizOverlay = dom.window.document.querySelector("#quizOverlay");
    assert.equal(quizModal.hidden, true);
    assert.equal(quizOverlay.hidden, true);
    assert.ok(!dom.window.document.body.classList.contains("quiz-open"));

    const select = dom.window.document.querySelector("#librarySelect");
    select.value = "doc-1";
    select.dispatchEvent(new dom.window.Event("change", { bubbles: true }));

    const primeBtn = dom.window.document.querySelector("#libraryActionBtn");
    const quizBtn = dom.window.document.querySelector("#libraryQuizBtn");
    primeBtn.click();
    await waitFor(() => quizBtn.disabled === false);

    quizBtn.click();
    await waitFor(() => dom.window.document.querySelectorAll(".quiz-question").length === 5);
    assert.equal(quizModal.hidden, false);

    dom.window.document.querySelector("#quizCloseBtn").click();
    await waitFor(() => quizModal.hidden === true);
    assert.equal(quizOverlay.hidden, true);
    assert.ok(!dom.window.document.body.classList.contains("quiz-open"));
    assert.equal(dom.window.document.querySelectorAll(".quiz-question").length, 0);

    quizBtn.click();
    await waitFor(() => dom.window.document.querySelectorAll(".quiz-question").length === 5);
    completeQuizSet(dom.window, 0);
    await waitFor(() => dom.window.document.querySelector("#quizGenerateMoreBtn").hidden === false);
    dom.window.document.querySelector("#quizGenerateMoreBtn").click();
    await waitFor(() => dom.window.document.querySelector("#quizInterruptBtn").hidden === false);
    interruptShouldFail = false;
    dom.window.document.querySelector("#quizInterruptBtn").click();
    await waitFor(() => quizModal.hidden === true);
    assert.ok(fetchCalls.some(url => url.startsWith("/api/library/quiz/interrupt")));
    if (delayedQuizSuccess) delayedQuizSuccess.resolve();

    quizBtn.click();
    await waitFor(() => dom.window.document.querySelectorAll(".quiz-question").length === 5);
    completeQuizSet(dom.window, 0);
    await waitFor(() => dom.window.document.querySelector("#quizGenerateMoreBtn").hidden === false);
    dom.window.document.querySelector("#quizGenerateMoreBtn").click();
    await waitFor(() => dom.window.document.querySelector("#quizInterruptBtn").hidden === false);
    interruptShouldFail = true;
    dom.window.document.querySelector("#quizInterruptBtn").click();
    const quizError = dom.window.document.querySelector("#quizError");
    const quizErrorText = dom.window.document.querySelector("#quizErrorText");
    await waitFor(() => quizError.hidden === false);
    assert.equal(quizModal.hidden, false);
    assert.equal(quizOverlay.hidden, false);
    assert.ok((quizErrorText.textContent || "").includes("Quiz interrupt failed."));
    dom.window.document.querySelector("#quizCloseBtn").click();
    await waitFor(() => quizModal.hidden === true);
    if (delayedQuizFailure) delayedQuizFailure.resolve();
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
    if (originalAlert) {
      globalAny.alert = originalAlert;
    } else {
      delete globalAny.alert;
    }
  }
});
