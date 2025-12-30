/**
 * Tests lecture quiz generation endpoint behavior.
 *
 * Used by: `npm test` to validate quiz UI flows and scoring.
 *
 * Key exports: none (test-only file).
 *
 * Key coverage:
 * - Quiz generation request/response flow, delayed responses, and answer scoring.
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

test("quiz on lecture button, flow, and scoring", async () => {
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
  let alertMessage = "";
  let quizCallCount = 0;
  let delayedQuizResponse = null;

  const quizBatch1 = {
    lectureId: "doc-1",
    lectureTitle: "Lecture One",
    setSize: 5,
    questions: [
      {
        id: "q1",
        stem: "What is the primary concept described in section one?",
        choices: [
          { id: "A", text: "Alpha" },
          { id: "B", text: "Beta" },
          { id: "C", text: "Gamma" },
          { id: "D", text: "Delta" },
        ],
        answer: "A",
        rationale: "Section one highlights Alpha as primary.",
        tags: ["intro", "overview"],
        difficulty: "easy",
        references: ["Slide 1 - Intro"],
      },
      {
        id: "q2",
        stem: "Which item best matches the lecture's definition?",
        choices: [
          { id: "A", text: "Option A" },
          { id: "B", text: "Option B" },
          { id: "C", text: "Option C" },
          { id: "D", text: "Option D" },
        ],
        answer: "B",
        rationale: "The lecture defines it as Option B.",
        tags: ["definition"],
        difficulty: "easy",
        references: ["Slide 2 - Definition"],
      },
      {
        id: "q3",
        stem: "Which mechanism is emphasized?",
        choices: [
          { id: "A", text: "Mechanism A" },
          { id: "B", text: "Mechanism B" },
          { id: "C", text: "Mechanism C" },
          { id: "D", text: "Mechanism D" },
        ],
        answer: "C",
        rationale: "Mechanism C is emphasized in the lecture.",
        tags: ["mechanism"],
        difficulty: "medium",
        references: ["Slide 3 - Mechanism"],
      },
      {
        id: "q4",
        stem: "Which statement is correct per the lecture?",
        choices: [
          { id: "A", text: "Statement A" },
          { id: "B", text: "Statement B" },
          { id: "C", text: "Statement C" },
          { id: "D", text: "Statement D" },
        ],
        answer: "D",
        rationale: "Statement D matches the lecture.",
        tags: ["summary"],
        difficulty: "medium",
        references: ["Slide 4 - Summary"],
      },
      {
        id: "q5",
        stem: "What is the key takeaway?",
        choices: [
          { id: "A", text: "Takeaway A" },
          { id: "B", text: "Takeaway B" },
          { id: "C", text: "Takeaway C" },
          { id: "D", text: "Takeaway D" },
        ],
        answer: "A",
        rationale: "Takeaway A is highlighted.",
        tags: ["takeaway"],
        difficulty: "easy",
        references: ["Slide 5 - Takeaways"],
      },
    ],
  };

  const quizBatch2 = {
    lectureId: "doc-1",
    lectureTitle: "Lecture One",
    setSize: 5,
    questions: [
      {
        id: "q6",
        stem: "Follow-up question one?",
        choices: [
          { id: "A", text: "Choice A" },
          { id: "B", text: "Choice B" },
          { id: "C", text: "Choice C" },
          { id: "D", text: "Choice D" },
        ],
        answer: "A",
        rationale: "Choice A is correct.",
        tags: ["follow-up"],
        difficulty: "easy",
        references: ["Slide 6 - Follow-up"],
      },
      {
        id: "q7",
        stem: "Follow-up question two?",
        choices: [
          { id: "A", text: "Choice A" },
          { id: "B", text: "Choice B" },
          { id: "C", text: "Choice C" },
          { id: "D", text: "Choice D" },
        ],
        answer: "B",
        rationale: "Choice B is correct.",
        tags: ["follow-up"],
        difficulty: "medium",
        references: ["Slide 7 - Follow-up"],
      },
      {
        id: "q8",
        stem: "Follow-up question three?",
        choices: [
          { id: "A", text: "Choice A" },
          { id: "B", text: "Choice B" },
          { id: "C", text: "Choice C" },
          { id: "D", text: "Choice D" },
        ],
        answer: "C",
        rationale: "Choice C is correct.",
        tags: ["follow-up"],
        difficulty: "medium",
        references: ["Slide 8 - Follow-up"],
      },
      {
        id: "q9",
        stem: "Follow-up question four?",
        choices: [
          { id: "A", text: "Choice A" },
          { id: "B", text: "Choice B" },
          { id: "C", text: "Choice C" },
          { id: "D", text: "Choice D" },
        ],
        answer: "D",
        rationale: "Choice D is correct.",
        tags: ["follow-up"],
        difficulty: "hard",
        references: ["Slide 9 - Follow-up"],
      },
      {
        id: "q10",
        stem: "Follow-up question five?",
        choices: [
          { id: "A", text: "Choice A" },
          { id: "B", text: "Choice B" },
          { id: "C", text: "Choice C" },
          { id: "D", text: "Choice D" },
        ],
        answer: "A",
        rationale: "Choice A is correct.",
        tags: ["follow-up"],
        difficulty: "easy",
        references: ["Slide 10 - Follow-up"],
      },
    ],
  };

  globalAny.document = dom.window.document;
  globalAny.window = dom.window;
  globalAny.navigator = dom.window.navigator;
  globalAny.HTMLElement = dom.window.HTMLElement;
  globalAny.Node = dom.window.Node;
  const alertStub = (message) => {
    alertMessage = message;
  };
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
    if (url.startsWith("/api/library/quiz")) {
      quizCallCount += 1;
      if (quizCallCount === 1) {
        return makeJsonResponse({ ok: true, batch: quizBatch1 });
      }
      if (quizCallCount === 2) {
        delayedQuizResponse = makeDelayedJsonResponse({ ok: true, batch: quizBatch2 });
        return delayedQuizResponse;
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

    const actions = dom.window.document.querySelector(".sidebar-actions");
    const buttons = Array.from(actions.querySelectorAll("button"));
    assert.equal(buttons[0]?.id, "libraryActionBtn");
    assert.equal(buttons[1]?.id, "libraryQuizBtn");

    const quizBtn = dom.window.document.querySelector("#libraryQuizBtn");
    assert.equal(quizBtn.disabled, true);
    assert.equal(alertMessage, "");

    const select = dom.window.document.querySelector("#librarySelect");
    select.value = "doc-1";
    select.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
    assert.equal(quizBtn.disabled, true);

    const primeBtn = dom.window.document.querySelector("#libraryActionBtn");
    primeBtn.click();
    await waitFor(() => quizBtn.disabled === false);

    quizBtn.click();
    await waitFor(() => dom.window.document.querySelectorAll(".quiz-question").length === 5);

    const ingestIndex = fetchCalls.findIndex(url => url.includes("/api/library/ingest"));
    const quizIndex = fetchCalls.findIndex(url => url.includes("/api/library/quiz"));
    assert.ok(ingestIndex !== -1);
    assert.ok(quizIndex !== -1);
    assert.ok(ingestIndex < quizIndex);

    const wrongOption = dom.window.document.querySelector(
      ".quiz-option[data-quiz-set-index=\"0\"][data-quiz-question-index=\"0\"][data-choice-id=\"B\"]",
    );
    wrongOption.click();
    const correctOption = dom.window.document.querySelector(
      ".quiz-option[data-quiz-set-index=\"0\"][data-quiz-question-index=\"0\"][data-choice-id=\"A\"]",
    );
    const explanation = dom.window.document.querySelector(
      ".quiz-question[data-quiz-set-index=\"0\"][data-quiz-question-index=\"0\"] [data-role=\"quiz-explanation\"]",
    );
    assert.equal(explanation.hidden, true);
    assert.ok(!correctOption.classList.contains("is-correct"));
    assert.ok(wrongOption.classList.contains("is-incorrect"));

    correctOption.click();
    assert.equal(explanation.hidden, false);

    dom.window.document.querySelector(
      ".quiz-option[data-quiz-set-index=\"0\"][data-quiz-question-index=\"1\"][data-choice-id=\"B\"]",
    ).click();

    dom.window.document.querySelector(
      ".quiz-option[data-quiz-set-index=\"0\"][data-quiz-question-index=\"2\"][data-choice-id=\"A\"]",
    ).click();
    dom.window.document.querySelector(
      ".quiz-option[data-quiz-set-index=\"0\"][data-quiz-question-index=\"2\"][data-choice-id=\"C\"]",
    ).click();

    dom.window.document.querySelector(
      ".quiz-option[data-quiz-set-index=\"0\"][data-quiz-question-index=\"3\"][data-choice-id=\"D\"]",
    ).click();
    dom.window.document.querySelector(
      ".quiz-option[data-quiz-set-index=\"0\"][data-quiz-question-index=\"4\"][data-choice-id=\"A\"]",
    ).click();

    const results = dom.window.document.querySelector(
      ".quiz-set[data-quiz-set-index=\"0\"] [data-quiz-set-results]",
    );
    assert.equal(results.hidden, false);
    const setScore = results.querySelector("[data-role=\"quiz-set-score\"]").textContent;
    const overallScore = results.querySelector("[data-role=\"quiz-overall-score\"]").textContent;
    assert.ok(setScore.includes("3/5"));
    assert.ok(setScore.includes("60%"));
    assert.ok(overallScore.includes("3/5"));
    assert.ok(overallScore.includes("60%"));

    const generateMoreBtn = dom.window.document.querySelector("#quizGenerateMoreBtn");
    assert.equal(generateMoreBtn.hidden, false);

    generateMoreBtn.click();
    const loading = dom.window.document.querySelector("#quizLoading");
    const interruptBtn = dom.window.document.querySelector("#quizInterruptBtn");
    assert.equal(loading.hidden, false);
    assert.ok(loading.textContent.includes("Owen is Working"));
    assert.equal(interruptBtn.hidden, false);
    assert.equal(interruptBtn.disabled, false);
    assert.equal(generateMoreBtn.disabled, true);

    delayedQuizResponse.resolve();
    await waitFor(() => dom.window.document.querySelectorAll(".quiz-set").length === 2);
    assert.equal(loading.hidden, true);
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
