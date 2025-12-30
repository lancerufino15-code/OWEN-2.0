/**
 * OWEN chat UI controller (browser module).
 *
 * Used by: `public/index.html` via `<script type="module" src="/chat.js">`.
 *
 * Responsibilities:
 * - Bind UI elements, handle chat input/streaming, uploads, retrieval, and history.
 * - Render assistant responses with citation pills and table parsing.
 *
 * Key exports:
 * - Rendering helpers used by tests or other UI modules.
 *
 * Assumptions:
 * - Runs in a browser with DOM access; no server-side rendering.
 */
import { MULTIWORD_PREPOSITION_REGEX, isArticleOrPreposition } from "./articles_prepositions.js";
import {
  buildCitationMapFromSegments,
  buildCitationPill,
  getCitationDomain,
  normalizeCitedSources,
  replaceCitationPlaceholders,
  resolveCitationLinks,
  isLikelyValidCitationUrl,
} from "./citation_utils.mjs";
import { isTableLikeBlock, renderHtmlTable, renderMarkdownTable } from "./table_parser.mjs";
import { createHoverSwapRobot } from "./hover_swap_robot.js";

let fileInput = null;
let fileButton = null;
let fileNameDisplay = null;
let nameInput = null;
let bucketSelect = null;
let sendButton = null;
let uploadStatus = null;
let uploadPanel = null;
let machineUploadToggle = null;
let machineUploadBody = null;
let libraryPanel = null;
let libraryPanelBody = null;
let libraryToggle = null;
let libraryAdvancedToggle = null;
let libraryAdvancedBody = null;

let retrievePanel = null;
let retrievePanelBody = null;
let retrieveToggle = null;
let retrieveInput = null;
let retrieveBtn = null;
let retrieveStatus = null;
let retrieveResult = null;

let savedPlaceholder = null;
let savedList = null;
let newConvoBtn = null;
let historyPanel = null;
let historyToggle = null;
let historyList = null;
let conversationSearch = null;

let chatLog = null;
let chatForm = null;
let chatInput = null;
let chatSend = null;
let chatAttach = null;
let chatFileInput = null;
let attachStatus = null;
let insightsLog = null;
let hashtagCloud = null;
let metaDataPanel = null;
let metaUnlockInput = null;
let metaLockBtn = null;
let metaToggle = null;
let metaBody = null;
let metaUnlockError = null;
let metaPanelToggle = null;
let metaTopicsChart = null;
let metaUpdated = null;
let metaLectureSelect = null;
let metaError = null;
let modelButtons = [];
let themeToggle = null;
let sidebar = null;
let sidebarCollapseToggle = null;
let sidebarDrawerToggle = null;
let toolsToggle = null;
let toolsOverlay = null;
let composerResizeObserver = null;
let toolsDrawerListenersBound = false;
let quizEscapeListenerBound = false;
let fullDocToggle = null;
let librarySearchInput = null;
let librarySearchBtn = null;
let libraryResults = null;
let libraryStatus = null;
let libraryIndexBtn = null;
let libraryIngestBtn = null;
let librarySelection = null;
let librarySelect = null;
let libraryActionBtn = null;
let libraryQuizBtn = null;
let libraryQuizHelper = null;
let libraryRefreshBtn = null;
let libraryClearBtn = null;
let libraryManagerSearch = null;
let libraryManagerList = null;
let libraryManagerStatus = null;
let machinePanel = null;
let machinePanelBody = null;
let machineToggle = null;
let machineTxtPanel = null;
let machineTxtToggle = null;
let machineTxtBody = null;
let machineStudyPanel = null;
let machineStudyToggle = null;
let machineStudyBody = null;
let machineUnlockInput = null;
let machineLockBtn = null;
let machineLectureSelect = null;
let machineGenerateBtn = null;
let machineStatus = null;
let machineResult = null;
let machineTxtDetails = null;
let machineTxtInput = null;
let machineTxtButton = null;
let machineTxtName = null;
let machineUseLastToggle = null;
let machineStudyMode = null;
let machineStudyBtn = null;
let machineStudyStatus = null;
let machineStudyResult = null;
let ankiPdfInput = null;
let ankiPdfButton = null;
let ankiParseButton = null;
let ankiImagesInput = null;
let ankiImagesButton = null;
let ankiTranscriptInput = null;
let ankiTranscriptButton = null;
let ankiBoldmapInput = null;
let ankiBoldmapButton = null;
let ankiClassmateInput = null;
let ankiClassmateButton = null;
let ankiGenerateBtn = null;
let ankiGenerateStatus = null;
let ankiGenerateResult = null;
let ankiStatusBadge = null;
let ankiLoadedList = null;
let ankiPdfStatus = null;
let ankiSlidesDownload = null;
let ankiCardsDownload = null;
let sourcesDrawer = null;
let sourcesDrawerClose = null;
let sourcesOverlay = null;
let sourcesList = null;
let sourcesSearch = null;
let quizOverlay = null;
let quizModal = null;
let quizCloseBtn = null;
let quizLectureTitle = null;
let quizLoading = null;
let quizInterruptBtn = null;
let quizError = null;
let quizErrorText = null;
let quizRetryBtn = null;
let quizSets = null;
let quizGenerateMoreBtn = null;
let facultyPasscodeInput = null;
let facultyUnlockBtn = null;
let facultyLoginError = null;
let facultyLogoutBtn = null;
let publishStudyGuideBtn = null;
let publishStatus = null;
let publishResult = null;
let retrieveTypeSelect = null;

const resetBoundElements = () => {
  fileInput = null;
  fileButton = null;
  fileNameDisplay = null;
  nameInput = null;
  bucketSelect = null;
  sendButton = null;
  uploadStatus = null;
  uploadPanel = null;
  machineUploadToggle = null;
  machineUploadBody = null;
  libraryPanel = null;
  libraryPanelBody = null;
  libraryToggle = null;
  libraryAdvancedToggle = null;
  libraryAdvancedBody = null;
  retrievePanel = null;
  retrievePanelBody = null;
  retrieveToggle = null;
  retrieveTypeSelect = null;
  retrieveInput = null;
  retrieveBtn = null;
  retrieveStatus = null;
  retrieveResult = null;
  savedPlaceholder = null;
  savedList = null;
  newConvoBtn = null;
  historyPanel = null;
  historyToggle = null;
  historyList = null;
  conversationSearch = null;
  chatLog = null;
  chatForm = null;
  chatInput = null;
  chatSend = null;
  chatAttach = null;
  chatFileInput = null;
  attachStatus = null;
  insightsLog = null;
  hashtagCloud = null;
  metaDataPanel = null;
  metaUnlockInput = null;
  metaLockBtn = null;
  metaToggle = null;
  metaBody = null;
  metaUnlockError = null;
  metaPanelToggle = null;
  metaTopicsChart = null;
  metaUpdated = null;
  metaLectureSelect = null;
  metaError = null;
  modelButtons = [];
  themeToggle = null;
  sidebar = null;
  sidebarCollapseToggle = null;
  sidebarDrawerToggle = null;
  toolsToggle = null;
  toolsOverlay = null;
  if (composerResizeObserver) {
    composerResizeObserver.disconnect();
    composerResizeObserver = null;
  }
  fullDocToggle = null;
  librarySearchInput = null;
  librarySearchBtn = null;
  libraryResults = null;
  libraryStatus = null;
  libraryIndexBtn = null;
  libraryIngestBtn = null;
  librarySelection = null;
  librarySelect = null;
  libraryActionBtn = null;
  libraryQuizBtn = null;
  libraryQuizHelper = null;
  libraryRefreshBtn = null;
  libraryClearBtn = null;
  libraryManagerSearch = null;
  libraryManagerList = null;
  libraryManagerStatus = null;
  machinePanel = null;
  machinePanelBody = null;
  machineToggle = null;
  machineTxtPanel = null;
  machineTxtToggle = null;
  machineTxtBody = null;
  machineStudyPanel = null;
  machineStudyToggle = null;
  machineStudyBody = null;
  machineUnlockInput = null;
  machineLockBtn = null;
  machineLectureSelect = null;
  machineGenerateBtn = null;
  machineStatus = null;
  machineResult = null;
  machineTxtDetails = null;
  machineTxtInput = null;
  machineTxtButton = null;
  machineTxtName = null;
  machineUseLastToggle = null;
  machineStudyMode = null;
  machineStudyBtn = null;
  machineStudyStatus = null;
  machineStudyResult = null;
  ankiPdfInput = null;
  ankiPdfButton = null;
  ankiParseButton = null;
  ankiImagesInput = null;
  ankiImagesButton = null;
  ankiTranscriptInput = null;
  ankiTranscriptButton = null;
  ankiBoldmapInput = null;
  ankiBoldmapButton = null;
  ankiClassmateInput = null;
  ankiClassmateButton = null;
  ankiGenerateBtn = null;
  ankiGenerateStatus = null;
  ankiGenerateResult = null;
  ankiStatusBadge = null;
  ankiLoadedList = null;
  ankiPdfStatus = null;
  ankiSlidesDownload = null;
  ankiCardsDownload = null;
  sourcesDrawer = null;
  sourcesDrawerClose = null;
  sourcesOverlay = null;
  sourcesList = null;
  sourcesSearch = null;
  quizOverlay = null;
  quizModal = null;
  quizCloseBtn = null;
  quizLectureTitle = null;
  quizLoading = null;
  quizInterruptBtn = null;
  quizError = null;
  quizErrorText = null;
  quizRetryBtn = null;
  quizSets = null;
  quizGenerateMoreBtn = null;
  facultyPasscodeInput = null;
  facultyUnlockBtn = null;
  facultyLoginError = null;
  facultyLogoutBtn = null;
  publishStudyGuideBtn = null;
  publishStatus = null;
  publishResult = null;
};

function queryByActionOrId(root, action, idSelector) {
  if (!root) return null;
  return root.querySelector(`[data-action="${action}"]`) || root.querySelector(idSelector);
}

function bindOnce(el, eventName, handler, key) {
  if (!el) return;
  const attr = `data-bound-${key || eventName}`;
  if (el.getAttribute(attr) === "true") return;
  el.setAttribute(attr, "true");
  el.addEventListener(eventName, handler);
}

const bindCommonElements = (root = document) => {
  themeToggle = root.querySelector("#theme-toggle");
  modelButtons = Array.from(root.querySelectorAll(".model-btn"));
};

const bindStudentElements = (root = document) => {
  bindCommonElements(root);
  sidebar = root.querySelector(".student-sidebar");
  sidebarCollapseToggle = root.querySelector("#sidebarCollapseToggle");
  sidebarDrawerToggle = root.querySelector("#sidebarDrawerToggle");
  toolsToggle = root.querySelector("#toolsToggle");
  toolsOverlay = root.querySelector("#toolsOverlay");
  libraryPanel = root.querySelector("#libraryPanel");
  libraryPanelBody = root.querySelector("#libraryPanelBody");
  libraryToggle = root.querySelector("#libraryToggle");
  libraryAdvancedToggle = root.querySelector("#libraryAdvancedToggle");
  libraryAdvancedBody = root.querySelector("#libraryAdvancedBody");
  retrievePanel = root.querySelector("#retrievePanel");
  retrievePanelBody = root.querySelector("#retrievePanelBody");
  retrieveToggle = root.querySelector("#retrieveToggle");
  retrieveTypeSelect = root.querySelector("#retrieveType");
  retrieveInput = root.querySelector("#retrieve-input");
  retrieveBtn = root.querySelector("#retrieve-btn");
  retrieveStatus = root.querySelector("#retrieve-status");
  retrieveResult = root.querySelector("#retrieve-result");
  savedPlaceholder = root.querySelector("#saved-placeholder");
  savedList = root.querySelector("#saved-list");
  newConvoBtn = root.querySelector("#new-convo");
  historyPanel = root.querySelector("#historyPanel");
  historyToggle = root.querySelector("#historyToggle");
  historyList = root.querySelector("#historyList");
  conversationSearch = root.querySelector("#conversation-search");
  chatLog = root.querySelector("#chat-log");
  chatForm = root.querySelector("#chat-form");
  chatInput = root.querySelector("#chat-input");
  chatSend = root.querySelector("#chat-send");
  chatAttach = root.querySelector("#chat-attach");
  chatFileInput = root.querySelector("#chat-file-input");
  attachStatus = root.querySelector("#attach-status");
  sourcesDrawer = root.querySelector("#sourcesDrawer");
  sourcesDrawerClose = root.querySelector("#sourcesDrawerClose");
  sourcesOverlay = root.querySelector("#sourcesOverlay");
  sourcesList = root.querySelector("#sourcesList");
  sourcesSearch = root.querySelector("#sourcesSearch");
  libraryStatus = root.querySelector("#library-status");
  libraryIndexBtn = root.querySelector("#library-index-btn");
  libraryIngestBtn = root.querySelector("#library-ingest-btn");
  librarySelect = root.querySelector("#librarySelect");
  libraryActionBtn = root.querySelector("#libraryActionBtn");
  libraryQuizBtn = root.querySelector("#libraryQuizBtn");
  libraryQuizHelper = root.querySelector("#libraryQuizHelper");
  libraryRefreshBtn = root.querySelector("#libraryRefresh");
  libraryClearBtn = root.querySelector("#libraryClear");
  quizOverlay = root.querySelector("#quizOverlay");
  quizModal = root.querySelector("#quizModal");
  quizCloseBtn = root.querySelector("#quizCloseBtn");
  quizLectureTitle = root.querySelector("#quizLectureTitle");
  quizLoading = root.querySelector("#quizLoading");
  quizInterruptBtn = root.querySelector("#quizInterruptBtn");
  quizError = root.querySelector("#quizError");
  quizErrorText = root.querySelector("#quizErrorText");
  quizRetryBtn = root.querySelector("#quizRetryBtn");
  quizSets = root.querySelector("#quizSets");
  quizGenerateMoreBtn = root.querySelector("#quizGenerateMoreBtn");
};

const bindFacultyElements = (root = document) => {
  bindCommonElements(root);
  facultyPasscodeInput = root.querySelector("#facultyPasscodeInput");
  facultyUnlockBtn = root.querySelector("#facultyUnlockBtn");
  facultyLoginError = root.querySelector("#facultyLoginError");
  facultyLogoutBtn = root.querySelector("#facultyLogout");
  publishStudyGuideBtn = queryByActionOrId(root, "publish-study-guide", "#publishStudyGuideBtn");
  publishStatus = root.querySelector("#publishStatus");
  publishResult = root.querySelector("#publishResult");
  fileInput = root.querySelector("#file-input");
  fileButton = root.querySelector("#file-button");
  fileNameDisplay = root.querySelector("#file-name");
  nameInput = root.querySelector("#name-input");
  bucketSelect = root.querySelector("#bucket-select");
  sendButton = root.querySelector("#send-button");
  uploadStatus = root.querySelector("#upload-status");
  fullDocToggle = root.querySelector("#full-doc-toggle");
  librarySelect = root.querySelector("#librarySelect");
  libraryActionBtn = root.querySelector("#libraryActionBtn");
  libraryStatus = root.querySelector("#library-status");
  libraryIndexBtn = root.querySelector("#library-index-btn");
  libraryIngestBtn = root.querySelector("#library-ingest-btn");
  libraryManagerSearch = root.querySelector("#libraryManagerSearch");
  libraryManagerList = root.querySelector("#libraryManagerList");
  libraryManagerStatus = root.querySelector("#libraryManagerStatus");
  machineLectureSelect = root.querySelector("#machineLectureSelect");
  machineGenerateBtn = root.querySelector("#machineGenerateBtn");
  machineStatus = root.querySelector("#machineStatus");
  machineResult = root.querySelector("#machineResult");
  machineTxtDetails = root.querySelector("#machineTxtDetails");
  machineTxtInput = root.querySelector("#machineTxtInput");
  machineTxtButton = root.querySelector("#machineTxtButton");
  machineTxtName = root.querySelector("#machineTxtName");
  machineUseLastToggle = root.querySelector("#machineUseLastTxt");
  machineStudyMode = root.querySelector("#machineStudyMode");
  machineStudyBtn = queryByActionOrId(root, "generate-study-guide", "#machineStudyBtn");
  machineStudyStatus = root.querySelector("#machineStudyStatus");
  machineStudyResult = root.querySelector("#machineStudyResult");
  ankiPdfInput = root.querySelector("#ankiPdfInput");
  ankiPdfButton = root.querySelector("#ankiPdfButton");
  ankiParseButton = root.querySelector("#ankiParseButton");
  ankiImagesInput = root.querySelector("#ankiImagesInput");
  ankiImagesButton = root.querySelector("#ankiImagesButton");
  ankiTranscriptInput = root.querySelector("#ankiTranscriptInput");
  ankiTranscriptButton = root.querySelector("#ankiTranscriptButton");
  ankiBoldmapInput = root.querySelector("#ankiBoldmapInput");
  ankiBoldmapButton = root.querySelector("#ankiBoldmapButton");
  ankiClassmateInput = root.querySelector("#ankiClassmateInput");
  ankiClassmateButton = root.querySelector("#ankiClassmateButton");
  ankiGenerateBtn = root.querySelector("#ankiGenerateBtn");
  ankiGenerateStatus = root.querySelector("#ankiGenerateStatus");
  ankiGenerateResult = root.querySelector("#ankiGenerateResult");
  ankiStatusBadge = root.querySelector("#ankiStatusBadge");
  ankiLoadedList = root.querySelector("#ankiLoadedList");
  ankiPdfStatus = root.querySelector("#ankiPdfStatus");
  ankiSlidesDownload = root.querySelector("#ankiSlidesDownload");
  ankiCardsDownload = root.querySelector("#ankiCardsDownload");
  metaBody = root.querySelector("#metaBody");
  metaUpdated = root.querySelector("#metaUpdated");
};

const attachmentUrlMap = window.__attachmentUrlMap || (window.__attachmentUrlMap = {});
const attachmentSummaries = window.__attachmentSummaries || (window.__attachmentSummaries = {});
const documentSummaryTasks = window.__documentSummaryTasks || (window.__documentSummaryTasks = new Set());
const attachments = [];
const conversation = [];
const responseState = new WeakMap();
let activeSourcesCard = null;
let activeSourceId = null;
let chatStreaming = false;
let activeOwenAbortController = null;
let activeOwenThinkingBubble = null;
let activeOwenStreamReader = null;
let libraryPendingDeleteId = null;
const libraryCopyFeedbackTimers = new WeakMap();
const TABLE_WRAP_PREF_KEY = "owen.tableWrapDefault";
const RESPONSE_VIEW_KEY = "owen.responseView";
const CITATION_MODE_KEY = "owen.citationMode";
const SIDEBAR_COLLAPSE_KEY = "owen.sidebarCollapsed";
const META_PANEL_VISIBLE_KEY = "owen.metaPanelVisible";
const MSG_RENDER_TOKEN = {};
let currentModel = "gpt-5-mini";
const THEME_KEY = "owen-theme";
const MODEL_STORAGE_KEY = "owen.model";
const attachmentUploads = new Set();
const CONVO_STORAGE_KEY = "owen.conversations";
const LAST_CONVO_KEY = "owen.activeConversationId";
const LEGACY_LAST_CONVO_KEY = "owen.conversations.active";
const CONVO_INDEX_KEY = "owen.conversationsIndex";
const CONVO_PREFIX = "owen.conversation.";
const CONVO_REMOTE_ENDPOINT = "/api/conversations";
const CONVO_REMOTE_SYNC_DEBOUNCE_MS = 900;
const CONVO_REMOTE_INDEX_TTL_MS = 30_000;
const CONVO_REMOTE_WARN_KEY = "owen.conversations.remoteWarned";
const TITLE_OVERRIDE_KEY = "owen.conversationTitleOverrides";
const MAX_STORED_MESSAGES = 400;
const MAX_STORED_CHARS = 600_000;
const MAX_MESSAGE_CHARS = 120_000;
const MAX_STORED_SOURCES = 80;
const MAX_STORED_CITATIONS = 120;
const MAX_STORED_SEGMENTS = 1200;

const LIGHT_BULB_ICON = `
  <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M8 9.5a4 4 0 1 1 8 0c0 1.6-.7 2.8-1.8 3.8-.5.5-.8 1.1-1 1.7h-2.4c-.2-.6-.5-1.2-1-1.7C8.7 12.3 8 11.1 8 9.5z" />
    <path d="M9.5 18h5" />
    <path d="M10.5 21h3" />
  </svg>
`.trim();

const JAGGED_MOON_ICON = `
  <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
  </svg>
`.trim();

const INTERRUPT_ICON = `
  <svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M4 4l8 8M12 4l-8 8" />
  </svg>
`.trim();

function setChatBusy(isBusy) {
  chatStreaming = Boolean(isBusy);
  if (chatInput) chatInput.disabled = chatStreaming;
  if (chatSend) chatSend.disabled = chatStreaming;
}

function createAbortError(message = "Interrupted") {
  try {
    return new DOMException(message, "AbortError");
  } catch {
    const err = new Error(message);
    err.name = "AbortError";
    return err;
  }
}

function isAbortError(error) {
  return error?.name === "AbortError";
}

function throwIfAborted(signal) {
  if (signal?.aborted) {
    throw signal.reason || createAbortError();
  }
}

function registerOwenAbortController(controller, bubble) {
  if (!controller) return null;
  activeOwenAbortController = controller;
  activeOwenThinkingBubble = bubble || null;
  activeOwenStreamReader = null;
  return controller;
}

function clearOwenAbortController(controller) {
  if (!controller) return;
  if (activeOwenAbortController !== controller) return;
  activeOwenAbortController = null;
  activeOwenThinkingBubble = null;
  activeOwenStreamReader = null;
}

function clearThinkingBubble(bubble, { removeEmpty = true } = {}) {
  if (!bubble || !bubble.isConnected) return;
  cancelTypewriter(bubble);
  bubble.classList.remove("thinking", "scaffold");
  const status = bubble.querySelector(".bubble-status");
  if (status) status.remove();
  const loading = bubble.querySelector(".loading-bar");
  if (loading) loading.remove();
  const textNode = bubble.querySelector(".bubble-text");
  const hasContent = Boolean(textNode && textNode.textContent && textNode.textContent.trim());
  if (removeEmpty && !hasContent) {
    bubble.remove();
  }
}

function interruptActiveThinking() {
  const controller = activeOwenAbortController;
  if (controller && !controller.signal.aborted) {
    controller.abort();
  }
  if (activeOwenStreamReader?.cancel) {
    try {
      activeOwenStreamReader.cancel();
    } catch {}
  }
  if (activePdfSession) {
    activePdfSession = { ...activePdfSession, cancelled: true };
    persistOcrSession(activePdfSession);
  }
  if (activeOwenThinkingBubble && activeOwenThinkingBubble.isConnected) {
    clearThinkingBubble(activeOwenThinkingBubble);
  }
  setChatBusy(false);
  clearOwenAbortController(controller);
}
const TRUNCATION_NOTICE = " … (truncated for storage)";
const OCR_PAGE_LIMIT = 15;
const PDFJS_VERSION = "4.10.38";
const PDF_RENDER_MAX_DIMENSION = 1600;
const PDF_JPEG_QUALITY = 0.75;
const ANKI_IMAGE_MAX_DIMENSION = 1600;
const ANKI_MAX_PAGES = 80;
const ANKI_MAX_IMAGES = 80;
const ANKI_PDF_ERROR = "Unsupported file type. Please upload a PDF (.pdf).";
const ANKI_IMAGES_ERROR = "Unsupported file type. Please upload JPG/PNG images.";
const ANKI_TRANSCRIPT_ERROR = "Unsupported file type. Please upload a TXT (.txt).";
const OCR_CONCURRENCY = 2;
const EARLY_OCR_READY_COUNT = 3;
const OCR_MAX_RETRIES = 3;
const OCR_SAVE_BATCH = 5;
const OCR_SESSION_KEY = "owen.ocr_session";
const DEST_STORAGE_KEY = "owen.uploadDestination";
const STUDY_GUIDE_STATE_KEY = "owen.studyGuideState";
const STUDY_GUIDE_PUBLISH_STATE_KEY = "owen.studyGuidePublishState";
const DESTINATIONS = [
  { label: "Anki Decks", value: "anki_decks" },
  { label: "Study Guides", value: "study_guides" },
  { label: "Library", value: "library" },
];
const CHAT_BUILD_ID = "20250222-1";
const TYPEWRITER_SPEED_MS = 28;
const TYPEWRITER_DEBOUNCE_MS = 90;
const CITATION_PLACEHOLDER_PREFIX = "__OWEN_CITE__";
const SOURCES_UNAVAILABLE_NOTE = "Sources unavailable for this response; please retry.";
const typewriterControllers = new Map();
const messageState = new Map();
const DEBUG_STREAMING = (() => {
  try {
    return localStorage.getItem("DEBUG_STREAMING") === "1" || localStorage.getItem("DEBUG_LECTURE_ASK") === "1";
  } catch {
    return false;
  }
})();
const DEBUG_LECTURE_UI = (() => {
  try {
    return localStorage.getItem("DEBUG_LECTURE_UI") === "1";
  } catch {
    return false;
  }
})();
const DEBUG_RENDER = (() => {
  try {
    return localStorage.getItem("DEBUG_RENDER") === "1";
  } catch {
    return false;
  }
})();
const DEBUG_TABLE_RAW = (() => {
  try {
    return localStorage.getItem("DEBUG_TABLE_RAW") === "1";
  } catch {
    return false;
  }
})();
const DEBUG_PERSIST = (() => {
  try {
    return localStorage.getItem("DEBUG_PERSIST") === "1";
  } catch {
    return false;
  }
})();

console.log("CHAT_BUILD_ID", CHAT_BUILD_ID);

function runRenderFallbackSelfTest() {
  if (!DEBUG_STREAMING) return;
  const bubble = document.createElement("div");
  bubble.className = "bubble assistant";
  bubble.dataset.msgId = "selftest";
  messageState.set("selftest", { streamText: "streamed text", lastAnswer: "" });
  renderFinalAssistant(bubble, "", new Map(), { debug: { requestId: "selftest" } });
  const rendered = bubble.querySelector(".bubble-text")?.textContent || "";
  console.log("[SELFTEST] render fallback", { renderedLen: rendered.length, rendered });
}

function runEmptyFinalRegressionTest() {
  if (!DEBUG_STREAMING) return;
  const bubble = document.createElement("div");
  bubble.className = "bubble assistant";
  bubble.dataset.msgId = "regression";
  messageState.set("regression", { streamText: "Hello world streamed", lastAnswer: "" });
  renderFinalAssistant(bubble, "", new Map(), {});
  const rendered = bubble.querySelector(".bubble-text")?.textContent || "";
  console.log("[SELFTEST] empty-final regression", { rendered });
}
let storedConversations = loadStoredConversations();
let titleOverrides = loadTitleOverrides();
let activeConversationId = loadActiveConversationId() || generateConversationId();
let activeConversationMeta = {
  id: activeConversationId,
  title: "Conversation",
  createdAt: Date.now(),
  updatedAt: Date.now(),
  selectedDocId: null,
  selectedDocTitle: "",
};
let activeRoute = "landing";
let openConversationMenuId = null;
let openExportMenuId = null;
let editingConversationId = null;
let pendingDeleteId = null;
let lastRemoteIndexFetch = 0;
let conversationLoadToken = 0;
const pendingConversationSync = new Map();
let conversationSyncTimer = null;
const warnedConversations = new Set();
try {
  const raw = safeGetItem(CONVO_REMOTE_WARN_KEY);
  const parsed = raw ? JSON.parse(raw) : [];
  if (Array.isArray(parsed)) {
    parsed.forEach(id => warnedConversations.add(id));
  }
} catch {}
const pendingMathNodes = [];
let katexIntervalStarted = false;
let metaDataUnlocked = false;
let metaDataExpanded = false;
let metaAnalyticsController = null;
let metaAnalyticsRequestId = 0;
let oneShotFile = null;
let oneShotPreviewUrl = "";
let activePdfSession = null;
let fullDocumentMode = true;
let activeLibraryDoc = null;
let libraryListItems = [];
let librarySearchResults = [];
let libraryPrimedDocId = null;
let libraryAdvancedOpen = false;
let quizSession = null;
let quizGenerating = false;
let quizAbortController = null;
let quizInterrupting = false;
let quizLastActiveElement = null;
let activeMachineDoc = null;
let machineGenerating = false;
let lastMachineTxt = null;
let machineStudyGenerating = false;
let lastStudyGuideOutput = null;
let lastPublishResult = null;
let publishInFlight = false;
let machineUnlocked = false;
let machineUnlockErrorTimer = null;
let ankiPdfFile = null;
let ankiImageFiles = [];
let ankiImagesSource = "";
let ankiTranscriptFile = null;
let ankiBoldmapFile = null;
let ankiClassmateFile = null;
let ankiGenerating = false;
let ankiPdfConverting = false;
let ankiGenerateAbortController = null;
let ankiGenerateRequestSeq = 0;
let ankiGenerateRequestId = "";
const ANKI_JOB_STATES = {
  EMPTY: "EMPTY",
  PDF_LOADED: "PDF_LOADED",
  PARSING: "PARSING",
  IMAGES_READY: "IMAGES_READY",
  TRANSCRIPT_READY: "TRANSCRIPT_READY",
  READY_TO_GENERATE: "READY_TO_GENERATE",
  GENERATING: "GENERATING",
  COMPLETE: "COMPLETE",
  ERROR: "ERROR",
};
let ankiJobState = ANKI_JOB_STATES.EMPTY;
let ankiLastStableState = ANKI_JOB_STATES.EMPTY;
let ankiJobError = "";
let ankiJobRecovery = "";
let ankiSlidesZipUrl = "";
let ankiSlidesZipName = "slides.zip";
let ankiSlidesOcrText = "";
let ankiManifest = null;
let ankiCardsDownloadUrl = "";
const FOLLOWUP_REGEX = /\b(?:yes|no|yeah|yep|correct|right|exactly|that|those|them|it|same|still|more|continue|go on|as before|as above|what about|how about|clarify|elaborate|thanks|please do)\b/i;
const STOPWORDS = new Set([
  "this", "that", "with", "from", "have", "will", "need", "more", "info", "what",
  "when", "where", "which", "would", "could", "should", "into", "about", "please", "tell",
  "year", "years", "data"
]);
const SECTION_LABEL_WHITELIST = ["Response", "Summary", "Diagnosis", "Labs", "Treatment", "Warnings", "Plan", "Next steps"];
const SECTION_LABEL_WHITELIST_LOWER = SECTION_LABEL_WHITELIST.map(label => label.toLowerCase());

const IMAGE_MODELS = new Set(["gpt-image-1", "gpt-image-1-mini", "dall-e-2", "dall-e-3"]);

const isImageModel = (model) => IMAGE_MODELS.has((model || "").toLowerCase());

let pdfjsLibPromise = null;
let jsZipPromise = null;
let storedOcrSession = loadStoredOcrSession();
initBucketSelect();

async function loadPdfJsLib() {
  if (pdfjsLibPromise) return pdfjsLibPromise;
  const moduleUrl = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}/pdf.min.mjs`;
  pdfjsLibPromise = import(moduleUrl)
    .then(mod => {
      const lib = mod?.default || mod;
      if (lib?.GlobalWorkerOptions) {
        lib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}/pdf.worker.min.mjs`;
      }
      return lib;
    })
    .catch(err => {
      pdfjsLibPromise = null;
      throw err;
    });
  return pdfjsLibPromise;
}

async function loadJsZip() {
  if (jsZipPromise) return jsZipPromise;
  const moduleUrl = "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";
  jsZipPromise = import(moduleUrl)
    .then(mod => mod?.default || mod)
    .catch(err => {
      jsZipPromise = null;
      throw err;
    });
  return jsZipPromise;
}

async function renderPdfPageToBlob(pdfjsLib, pdfDoc, pageNumber) {
  const page = await pdfDoc.getPage(pageNumber);
  const baseViewport = page.getViewport({ scale: 1 });
  const baseMax = Math.max(baseViewport.width, baseViewport.height);
  const scale = Math.max(0.5, Math.min(PDF_RENDER_MAX_DIMENSION / baseMax, 1.6));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.ceil(viewport.width));
  canvas.height = Math.max(1, Math.ceil(viewport.height));
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("Canvas context unavailable for PDF rendering.");
  }
  const renderTask = page.render({ canvasContext: ctx, viewport });
  await renderTask.promise;
  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob((b) => {
      if (b) resolve(b);
      else reject(new Error("PDF render to blob failed."));
    }, "image/jpeg", PDF_JPEG_QUALITY);
  });
  canvas.width = canvas.height = 0;
  page.cleanup?.();
  const size = blob?.size || 0;
  console.log("[PDF] Rendered page", {
    pageNumber,
    width: Math.round(viewport.width),
    height: Math.round(viewport.height),
    size,
  });
  if (size > 1_000_000) {
    console.warn("[PDF] Render size above 1MB, consider lowering quality", { pageNumber, size });
  }
  return { blob, width: Math.round(viewport.width), height: Math.round(viewport.height), size };
}

function loadImageFromBlob(blob) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const url = URL.createObjectURL(blob);
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Image load failed."));
    };
    img.src = url;
  });
}

async function convertImageFileToJpeg(file, filename) {
  const blob = file instanceof Blob ? file : new Blob([file], { type: file?.type || "application/octet-stream" });
  const img = await loadImageFromBlob(blob);
  const width = img.naturalWidth || img.width || 1;
  const height = img.naturalHeight || img.height || 1;
  const maxDim = Math.max(width, height);
  const scale = maxDim > ANKI_IMAGE_MAX_DIMENSION ? ANKI_IMAGE_MAX_DIMENSION / maxDim : 1;
  const targetWidth = Math.max(1, Math.round(width * scale));
  const targetHeight = Math.max(1, Math.round(height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = targetWidth;
  canvas.height = targetHeight;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("Canvas context unavailable for image conversion.");
  }
  ctx.drawImage(img, 0, 0, targetWidth, targetHeight);
  const outBlob = await new Promise((resolve, reject) => {
    canvas.toBlob((b) => {
      if (b) resolve(b);
      else reject(new Error("Image conversion failed."));
    }, "image/jpeg", PDF_JPEG_QUALITY);
  });
  canvas.width = canvas.height = 0;
  return new File([outBlob], filename, { type: "image/jpeg" });
}

async function computeSha256(arrayBuffer) {
  const digest = await crypto.subtle.digest("SHA-256", arrayBuffer);
  return Array.from(new Uint8Array(digest))
    .map(byte => byte.toString(16).padStart(2, "0"))
    .join("");
}

function isPdfFile(file) {
  const name = (file?.name || "").toLowerCase();
  const type = (file?.type || "").toLowerCase();
  return name.endsWith(".pdf") || type.includes("pdf");
}

function isAnkiPdfFile(file) {
  return Boolean(file) && isPdfFile(file);
}

function isAnkiImageFile(file) {
  if (!file) return false;
  const name = (file.name || "").toLowerCase();
  const type = (file.type || "").toLowerCase();
  if (type.includes("jpeg") || type.includes("jpg") || type.includes("png")) return true;
  return name.endsWith(".jpg") || name.endsWith(".jpeg") || name.endsWith(".png");
}

function isAnkiTranscriptFile(file) {
  if (!file) return false;
  const name = (file.name || "").toLowerCase();
  const type = (file.type || "").toLowerCase();
  if (type === "text/plain" || type.startsWith("text/plain")) return true;
  if (type === "" || type === "application/octet-stream") {
    return name.endsWith(".txt");
  }
  return name.endsWith(".txt");
}

function extractPageNumberFromFilename(name) {
  const lowered = (name || "").toLowerCase();
  const match = lowered.match(/(?:^|[^a-z0-9])p(\d{1,5})(?:[^a-z0-9]|$)/i);
  if (match) return Number.parseInt(match[1], 10);
  const fallback = lowered.match(/(\d{1,5})/);
  return fallback ? Number.parseInt(fallback[1], 10) : null;
}

async function normalizeAnkiImageFiles(files) {
  const entries = files.map((file, index) => ({
    file,
    index,
    pageHint: extractPageNumberFromFilename(file?.name || "") ?? Number.NaN,
  }));
  entries.sort((a, b) => {
    const aHint = Number.isFinite(a.pageHint) ? a.pageHint : Number.POSITIVE_INFINITY;
    const bHint = Number.isFinite(b.pageHint) ? b.pageHint : Number.POSITIVE_INFINITY;
    if (aHint !== bHint) return aHint - bHint;
    const nameA = (a.file?.name || "").toLowerCase();
    const nameB = (b.file?.name || "").toLowerCase();
    if (nameA && nameB && nameA !== nameB) return nameA.localeCompare(nameB);
    return a.index - b.index;
  });
  const normalized = [];
  for (let i = 0; i < entries.length; i += 1) {
    setAnkiPdfStatus("pending", `Normalizing image ${i + 1} / ${entries.length}`);
    const filename = `${i + 1}.jpg`;
    const converted = await convertImageFileToJpeg(entries[i].file, filename);
    normalized.push(converted);
  }
  return normalized;
}

function formatAnkiFileSize(bytes) {
  const size = Number(bytes);
  if (!Number.isFinite(size) || size <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = size;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  const rounded = value >= 10 || unitIndex === 0 ? Math.round(value) : Math.round(value * 10) / 10;
  return `${rounded} ${units[unitIndex]}`;
}

function formatAnkiFileTooltip(file) {
  const name = file?.name || "file";
  const type = file?.type || "unknown";
  const size = typeof file?.size === "number" ? formatAnkiFileSize(file.size) : "unknown size";
  return `${name} | ${type} | ${size}`;
}

function buildAnkiPdfError(file) {
  if (!file) return ANKI_PDF_ERROR;
  const name = file.name || "file";
  const type = file.type || "unknown";
  const hasPdfExt = name.toLowerCase().endsWith(".pdf");
  const hasPdfMime = (file.type || "").toLowerCase().includes("pdf");
  if (!hasPdfExt && !hasPdfMime) {
    return `PDF file rejected: ${name} has no .pdf extension and MIME type is ${type}.`;
  }
  if (!hasPdfExt) {
    return `PDF file rejected: ${name} has no .pdf extension.`;
  }
  if (!hasPdfMime) {
    return `PDF file rejected: ${name} has MIME type ${type}.`;
  }
  return ANKI_PDF_ERROR;
}

function buildAnkiImagesError(file) {
  if (!file) return ANKI_IMAGES_ERROR;
  const name = file.name || "file";
  const type = file.type || "unknown";
  return `Image file rejected: ${name} is not a JPG/PNG (${type}).`;
}

function buildAnkiTranscriptError(file) {
  if (!file) return ANKI_TRANSCRIPT_ERROR;
  const name = file.name || "file";
  const type = file.type || "unknown";
  return `Transcript rejected: expected .txt, got ${name} (${type}).`;
}

function createAnkiRequestId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  const rand = Math.random().toString(16).slice(2, 10);
  return `anki-${Date.now().toString(16)}-${rand}`;
}

function appendAnkiRequestId(message, requestId) {
  if (!requestId) return message;
  if ((message || "").includes("Request ID")) return message;
  return `${message} (Request ID: ${requestId})`;
}

function resolveAnkiResponseRequestId(payload, resp, fallback) {
  const headerId = resp?.headers?.get("x-request-id") || "";
  return payload?.requestId || headerId || fallback || "";
}

function isAnkiTypeErrorMessage(el) {
  const msg = (el?.textContent || "").toLowerCase();
  return (
    msg.includes("file type") ||
    msg.includes("not supported") ||
    msg.includes("try again with a pdf") ||
    msg.includes("only pdf") ||
    msg.includes("upload a pdf") ||
    msg.includes("upload a txt") ||
    msg.includes("upload jpg") ||
    msg.includes("pdf file rejected") ||
    msg.includes("image file rejected") ||
    msg.includes("transcript rejected") ||
    msg.includes("must be .pdf") ||
    msg.includes("must be .txt") ||
    msg.includes("jpg/png") ||
    msg.includes("slides are required") ||
    msg.includes("transcript is required")
  );
}

function clearAnkiGenerateError() {
  if (ankiJobState !== ANKI_JOB_STATES.ERROR) return;
  clearAnkiJobError();
  updateAnkiGenerateState();
}

async function promisePool(items, limit, iterator) {
  const results = [];
  const executing = [];
  for (const item of items) {
    const p = Promise.resolve().then(() => iterator(item));
    results.push(p);
    if (limit > 0) {
      const task = p.then(() => {
        const idx = executing.indexOf(task);
        if (idx >= 0) executing.splice(idx, 1);
      });
      executing.push(task);
      if (executing.length >= limit) {
        await Promise.race(executing);
      }
    }
  }
  return Promise.all(results);
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function wantsEntireDocument(prompt) {
  if (!prompt || typeof prompt !== "string") return false;
  return /\b(entire|whole|all)\s+(document|pdf|file|pages)\b/i.test(prompt) || /\bfull\s+(document|pdf)\b/i.test(prompt);
}

function normalizeMessageAttachment(att) {
  if (!att || typeof att !== "object") return null;
  const filename = typeof att.filename === "string" && att.filename.trim()
    ? att.filename.trim()
    : (typeof att.key === "string" && att.key.trim()) || "attachment";
  const url = typeof att.url === "string" ? att.url : undefined;
  const bucket = typeof att.bucket === "string" ? att.bucket : undefined;
  const key = typeof att.key === "string" ? att.key : undefined;
  const textKey = typeof att.textKey === "string" ? att.textKey : undefined;
  const ocrStatus = typeof att.ocrStatus === "string" ? att.ocrStatus : undefined;
  const fileId = typeof att.fileId === "string" ? att.fileId : undefined;
  const visionFileId = typeof att.visionFileId === "string" ? att.visionFileId : undefined;
  const ocrWarning = typeof att.ocrWarning === "string" ? att.ocrWarning : undefined;
  const visionWarning = typeof att.visionWarning === "string" ? att.visionWarning : undefined;
  const mimeType = typeof att.mimeType === "string" ? att.mimeType : undefined;
  let previewUrl = typeof att.previewUrl === "string" ? att.previewUrl : undefined;
  const looksLikeImage = isImageAttachment({ mimeType, filename, key, url, previewUrl });
  if (!previewUrl && looksLikeImage && url) {
    previewUrl = url;
  }
  return {
    filename,
    url,
    bucket,
    key,
    textKey,
    ocrStatus,
    fileId,
    visionFileId,
    ocrWarning,
    visionWarning,
    mimeType,
    previewUrl,
  };
}

function cloneMessageAttachments(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map(normalizeMessageAttachment)
    .filter(Boolean);
}

function isImageAttachment(att) {
  const mime = (att?.mimeType || "").toLowerCase();
  const name = (att?.filename || att?.key || "").toLowerCase();
  const url = (att?.url || "").toLowerCase();
  const preview = (att?.previewUrl || "").toLowerCase();
  if (mime.startsWith("image/")) return true;
  if (preview.startsWith("data:image")) return true;
  if (/\.(png|jpe?g|gif|webp|bmp|svg|heic|heif)$/i.test(name)) return true;
  if (/\.(png|jpe?g|gif|webp|bmp|svg|heic|heif)$/i.test(url)) return true;
  return false;
}

function buildFileReferencesFromAttachments(list) {
  return cloneMessageAttachments(list)
    .map(att => {
      const bucket = typeof att.bucket === "string" ? att.bucket.trim() : "";
      const key = typeof att.key === "string" ? att.key.trim() : "";
      if (!bucket || !key) return null;
      const displayName = typeof att.filename === "string" ? att.filename.trim() : "";
      const fileId = typeof att.fileId === "string" ? att.fileId.trim() : "";
      const visionFileId = typeof att.visionFileId === "string" ? att.visionFileId.trim() : "";
      const ref = {
        bucket,
        key,
        displayName: displayName || att.filename || key,
      };
      const textKey = typeof att.textKey === "string" ? att.textKey.trim() : "";
      if (textKey) {
        ref.textKey = textKey;
      }
      if (fileId) {
        ref.fileId = fileId;
      }
      if (visionFileId) {
        ref.visionFileId = visionFileId;
      }
      return ref;
    })
    .filter(Boolean);
}

function ensureBubbleTextNode(bubble, { reset = false } = {}) {
  if (!(bubble instanceof HTMLElement)) return null;
  if (reset) {
    bubble.innerHTML = "";
  }
  let node = bubble.querySelector(".bubble-text");
  if (!node) {
    node = document.createElement("div");
    node.className = "bubble-text";
    bubble.appendChild(node);
  }
  return node;
}

function ensureBubbleStatusNode(bubble) {
  if (!(bubble instanceof HTMLElement)) return null;
  let status = bubble.querySelector(".bubble-status");
  if (!status) {
    status = document.createElement("div");
    status.className = "bubble-status";
    const text = document.createElement("span");
    text.className = "bubble-status-text";
    status.appendChild(text);
    bubble.appendChild(status);
  }
  return status;
}

function safeReplaceNode(oldNode, newNode, context = "replace") {
  if (!oldNode || !newNode) return false;
  const parent = oldNode.parentNode;
  if (!parent) return false;
  if (!parent.contains(oldNode)) return false;
  try {
    if (typeof oldNode.replaceWith === "function") {
      oldNode.replaceWith(newNode);
    } else {
      parent.replaceChild(newNode, oldNode);
    }
    return true;
  } catch (err) {
    console.warn(`[UI:${context}] replace failed`, err);
    return false;
  }
}

function renderBubbleAttachments(list) {
  const sanitized = cloneMessageAttachments(list);
  if (!sanitized.length) return null;
  const wrapper = document.createElement("div");
  wrapper.className = "bubble-attachments";
  sanitized.forEach(att => {
    const isImage = isImageAttachment(att);
    const item = document.createElement(isImage ? "span" : att.url ? "a" : "span");
    item.className = "bubble-attachment";
    if (isImage) item.classList.add("image-only");
    const thumbUrl = isImage ? (att.previewUrl || att.url || att.dataUrl) : "";
    if (isImage && thumbUrl) {
      const img = document.createElement("img");
      img.src = thumbUrl;
      img.alt = att.filename || "image attachment";
      img.loading = "lazy";
      img.className = "bubble-attachment-thumb";
      item.appendChild(img);
    }
    const label = document.createElement("span");
    label.className = "bubble-attachment-label";
    label.textContent = att.filename ? `📎 ${att.filename}` : "Attachment";
    item.appendChild(label);
    if (att.url) {
      item.href = att.url;
      item.target = "_blank";
      item.rel = "noopener noreferrer";
    }
    wrapper.appendChild(item);
  });
  return wrapper;
}

const textFromParts = (parts) => {
  if (!parts) return "";
  if (typeof parts === "string") return parts;
  if (Array.isArray(parts)) return parts.map(textFromParts).join("");
  if (typeof parts === "object") {
    if (typeof parts.text === "string") return parts.text;
    if (Array.isArray(parts.content)) return parts.content.map(textFromParts).join("");
  }
  return "";
};

function escapeHTML(str) {
  return str.replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch] || ch));
}

function getKeyLeaf(key) {
  if (typeof key !== "string") return "";
  const parts = key.split("/").filter(Boolean);
  return parts.length ? parts[parts.length - 1] : key;
}

function cleanLectureLabel(raw) {
  if (typeof raw !== "string") return "";
  const trimmed = raw.trim();
  if (!trimmed) return "";
  let cleaned = trimmed.replace(/^\d+\s+/, "");
  cleaned = cleaned.replace(/^\d+[_-]+/, "");
  cleaned = cleaned.replace(/^[a-f0-9]{8,}[_-]+/i, "");
  cleaned = cleaned.replace(/\.pdf$/i, "");
  cleaned = cleaned.replace(/\.txt$/i, "");
  cleaned = cleaned.trim();
  return cleaned || trimmed;
}

function scrubLectureFilenamePrefix(raw) {
  if (typeof raw !== "string") return "";
  const trimmed = raw.trim();
  if (!trimmed) return "";
  const leaf = trimmed.split("/").pop() || trimmed;
  const cleaned = leaf.replace(/^[0-9]+[-_]+/, "").trim();
  return cleaned || leaf;
}

function getLectureLabelSource(item) {
  if (!item || typeof item !== "object") return "";
  if (typeof item.originalLectureTitle === "string" && item.originalLectureTitle.trim()) return item.originalLectureTitle.trim();
  if (typeof item.sourceFilename === "string" && item.sourceFilename.trim()) return item.sourceFilename.trim();
  if (typeof item.originalFilename === "string" && item.originalFilename.trim()) return item.originalFilename.trim();
  if (typeof item.filename === "string" && item.filename.trim()) return item.filename.trim();
  if (typeof item.title === "string" && item.title.trim()) return item.title.trim();
  if (typeof item.key === "string" && item.key.trim()) return getKeyLeaf(item.key).trim();
  if (typeof item.docId === "string" && item.docId.trim()) return item.docId.trim();
  return "";
}

function getLectureBaseLabel(item) {
  if (item && typeof item === "object" && typeof item.displayName === "string" && item.displayName.trim()) {
    return scrubLectureFilenamePrefix(item.displayName);
  }
  const source = getLectureLabelSource(item);
  const cleaned = cleanLectureLabel(source);
  if (cleaned) return cleaned;
  if (source) return source;
  if (item && typeof item === "object" && typeof item.docId === "string") return item.docId;
  return "";
}

function getLectureDisplayLabel(item) {
  if (item && typeof item.displayLabel === "string" && item.displayLabel.trim()) return item.displayLabel;
  return getLectureBaseLabel(item);
}

function getLectureTxtDisplayName(source) {
  const base = cleanLectureLabel(typeof source === "string" ? source : "") || (typeof source === "string" ? source.trim() : "");
  const safe = base.replace(/\.txt$/i, "").replace(/\.pdf$/i, "").trim() || "Lecture";
  return `${safe}.txt`;
}

function applyLectureDisplayLabels(items) {
  const seen = new Map();
  (items || []).forEach(item => {
    const base = getLectureBaseLabel(item) || (item?.docId ? String(item.docId) : "Lecture");
    const count = (seen.get(base) || 0) + 1;
    seen.set(base, count);
    const label = count === 1 ? base : `${base} (${count})`;
    if (item && typeof item === "object") {
      item.displayLabel = label;
    }
  });
}

function formatLibraryTimestamp(value) {
  if (!value) return "--";
  const parsed = new Date(value);
  if (!Number.isFinite(parsed.valueOf())) return String(value);
  return formatRelativeTime(parsed.getTime());
}

const USER_VISIBLE_LIBRARY_ARTIFACTS = [
  { key: "pdf", label: "PDF", downloadLabel: "Download PDF" },
  { key: "txt", label: "TXT", downloadLabel: "Download TXT" },
  { key: "guide", label: "Guide", downloadLabel: "Download Study Guide" },
];

function buildLibraryDownloadOptions(item) {
  const artifacts = item?.artifacts || {};
  const options = [];
  USER_VISIBLE_LIBRARY_ARTIFACTS.forEach(def => {
    if (artifacts?.[def.key]?.exists) {
      options.push({ value: def.key, label: def.downloadLabel });
    }
  });
  return options.map(opt => `<option value="${opt.value}">${opt.label}</option>`).join("");
}

function renderLibraryManagerList() {
  if (!libraryManagerList) return;
  const query = (libraryManagerSearch?.value || "").trim().toLowerCase();
  const items = (libraryListItems || []).filter(item => {
    if (!query) return true;
    const blob = [
      getLectureDisplayLabel(item),
      item?.title,
      item?.key,
      item?.bucket,
      item?.docId,
    ].filter(Boolean).join(" ").toLowerCase();
    return blob.includes(query);
  });

  if (!items.length) {
    const message = libraryListItems?.length
      ? "No lectures match this search."
      : "No lectures indexed yet.";
    libraryManagerList.innerHTML = `<div class="library-manager__empty">${escapeHTML(message)}</div>`;
  } else {
    libraryManagerList.innerHTML = items.map(item => {
      const label = getLectureDisplayLabel(item) || item?.title || item?.key || item?.docId || "Lecture";
      const docId = item?.docId || "";
      const updatedAt = formatLibraryTimestamp(item?.updatedAt || item?.uploaded);
      const accessCode = typeof item?.accessCode === "string" ? item.accessCode.trim() : "";
      const codeDisplay = accessCode ? escapeHTML(accessCode) : "&mdash;";
      const metaParts = [];
      if (updatedAt && updatedAt !== "--") {
        metaParts.push(`<span class="library-manager__meta-item">Updated ${escapeHTML(updatedAt)}</span>`);
      }
      const copyButton = accessCode
        ? `
          <button
            type="button"
            class="library-code-copy"
            data-library-action="copy-code"
            data-doc-id="${escapeHTML(docId)}"
            data-access-code="${escapeHTML(accessCode)}"
            aria-label="Copy access code"
            title="Copy access code"
          >
            <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <rect x="9" y="9" width="11" height="11" rx="2"></rect>
              <rect x="4" y="4" width="11" height="11" rx="2"></rect>
            </svg>
          </button>
        `
        : "";
      const codeMeta = `
        <span class="library-manager__meta-item library-manager__meta-code">
          <span>Code:</span>
          <span class="library-manager__code">${codeDisplay}</span>
          ${copyButton}
          <span class="library-code-feedback" role="status" aria-live="polite"></span>
        </span>
      `;
      if (metaParts.length) metaParts.push(`<span class="library-manager__meta-sep">&middot;</span>`);
      metaParts.push(codeMeta);
      const metaHtml = metaParts.length ? `<div class="library-manager__meta">${metaParts.join("")}</div>` : "";
      const downloadOptions = buildLibraryDownloadOptions(item);
      const downloadDisabled = !downloadOptions;
      const confirmOpen = libraryPendingDeleteId && docId && libraryPendingDeleteId === docId;
      const confirmHtml = confirmOpen ? `
        <div class="library-manager__confirm">
          <div>Type DELETE to confirm deleting "${escapeHTML(label)}".</div>
          <div class="library-manager__confirm-actions">
            <input class="sidebar-input library-manager__confirm-input" data-library-input="confirm-delete" data-doc-id="${escapeHTML(docId)}" placeholder="Type DELETE" />
            <button class="btn btn-danger" data-library-action="confirm-delete" data-doc-id="${escapeHTML(docId)}" disabled>Delete</button>
            <button class="btn btn-ghost" data-library-action="cancel-delete" data-doc-id="${escapeHTML(docId)}">Cancel</button>
          </div>
        </div>
      ` : "";
      return `
        <div class="library-manager__row" data-doc-id="${escapeHTML(docId)}">
          <div>
            <div class="library-manager__title">${escapeHTML(label)}</div>
            ${metaHtml}
          </div>
          <div class="library-manager__actions">
            <select class="sidebar-select library-manager__download" data-library-action="download" data-doc-id="${escapeHTML(docId)}" ${downloadDisabled ? "disabled" : ""}>
              <option value="">Download...</option>
              ${downloadOptions}
            </select>
            <button class="btn btn-danger" data-library-action="prompt-delete" data-doc-id="${escapeHTML(docId)}">Delete</button>
          </div>
          ${confirmHtml}
        </div>
      `.trim();
    }).join("");
  }

  if (libraryManagerStatus) {
    if (!libraryListItems?.length) {
      setStatus(libraryManagerStatus, "info", "No lectures indexed yet.");
    } else if (query) {
      setStatus(libraryManagerStatus, "info", `Showing ${items.length} of ${libraryListItems.length} lectures.`);
    } else {
      setStatus(libraryManagerStatus, "info", `${libraryListItems.length} lectures in library.`);
    }
  }
}

function triggerLibraryDownload(docId, type) {
  if (!docId || !type) return;
  const url = `/api/library/download?lectureId=${encodeURIComponent(docId)}&type=${encodeURIComponent(type)}`;
  const link = document.createElement("a");
  link.href = url;
  link.download = "";
  document.body.appendChild(link);
  link.click();
  link.remove();
}

async function deleteLibraryLecture(docId, label) {
  if (!docId) return;
  if (libraryManagerStatus) {
    setStatus(libraryManagerStatus, "pending", `Deleting ${label}...`);
  }
  try {
    const resp = await facultyFetch(
      `/api/library/lecture?lectureId=${encodeURIComponent(docId)}`,
      { method: "DELETE" },
      "library_delete",
    );
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok || payload?.ok === false) {
      const msg = payload?.error || payload?.message || "Delete failed.";
      const failureCount = Array.isArray(payload?.failures) ? payload.failures.length : 0;
      const exampleKey = failureCount ? payload?.failures?.[0]?.key : "";
      const detail = failureCount
        ? ` (${failureCount} artifact${failureCount === 1 ? "" : "s"} failed${exampleKey ? `, e.g. ${exampleKey}` : ""})`
        : "";
      throw new Error(`${msg}${detail}`);
    }
    libraryPendingDeleteId = null;
    await loadLibraryList();
    if (libraryManagerStatus) {
      setStatus(libraryManagerStatus, "success", `Deleted ${label}.`);
    }
  } catch (error) {
    if (libraryManagerStatus) {
      setStatus(libraryManagerStatus, "error", error instanceof Error ? error.message : "Delete failed.");
    }
    throw error;
  }
}

function showLibraryCodeFeedback(target, message, opts = {}) {
  if (!(target instanceof Element)) return;
  const row = target.closest(".library-manager__row");
  const feedback = row?.querySelector(".library-code-feedback");
  if (!feedback) return;
  feedback.textContent = message;
  feedback.classList.toggle("is-error", Boolean(opts.error));
  feedback.classList.add("is-visible");
  if (opts.copied) {
    target.setAttribute("data-copied", "true");
  } else {
    target.removeAttribute("data-copied");
  }
  const previous = libraryCopyFeedbackTimers.get(feedback);
  if (previous) clearTimeout(previous);
  const timeout = setTimeout(() => {
    feedback.classList.remove("is-visible");
    feedback.classList.remove("is-error");
    feedback.textContent = "";
    target.removeAttribute("data-copied");
  }, 1200);
  libraryCopyFeedbackTimers.set(feedback, timeout);
}

async function handleLibraryManagerClick(event) {
  const target = event.target instanceof Element ? event.target.closest("[data-library-action]") : null;
  if (!target) return;
  const action = target.getAttribute("data-library-action") || "";
  const docId = target.getAttribute("data-doc-id") || "";
  if (action === "copy-code") {
    const accessCode = target.getAttribute("data-access-code") || "";
    if (!accessCode) return;
    const ok = await copyTextToClipboard(accessCode);
    if (ok) {
      showLibraryCodeFeedback(target, "Copied", { copied: true });
    } else {
      showLibraryCodeFeedback(target, "Copy failed", { error: true });
    }
    return;
  }
  if (!docId) return;
  if (action === "prompt-delete") {
    libraryPendingDeleteId = docId;
    renderLibraryManagerList();
    return;
  }
  if (action === "cancel-delete") {
    libraryPendingDeleteId = null;
    renderLibraryManagerList();
    return;
  }
  if (action === "confirm-delete") {
    const row = target.closest(".library-manager__row");
    const input = row?.querySelector('[data-library-input="confirm-delete"]');
    const typed = (input?.value || "").trim().toUpperCase();
    if (typed !== "DELETE") {
      if (libraryManagerStatus) {
        setStatus(libraryManagerStatus, "error", "Type DELETE to confirm deletion.");
      }
      return;
    }
    const item = libraryListItems.find(entry => entry.docId === docId);
    const label = getLectureDisplayLabel(item) || item?.title || docId;
    deleteLibraryLecture(docId, label).catch(() => {});
  }
}

function handleLibraryManagerChange(event) {
  const target = event.target instanceof Element ? event.target : null;
  if (!target || !target.matches(".library-manager__download")) return;
  const docId = target.getAttribute("data-doc-id") || "";
  const type = target.value || "";
  if (!docId || !type) return;
  triggerLibraryDownload(docId, type);
  target.value = "";
}

function handleLibraryManagerInput(event) {
  const target = event.target instanceof Element ? event.target : null;
  if (!target || target.getAttribute("data-library-input") !== "confirm-delete") return;
  const row = target.closest(".library-manager__row");
  const confirmBtn = row?.querySelector('[data-library-action="confirm-delete"]');
  const typed = (target.value || "").trim().toUpperCase();
  if (confirmBtn) {
    confirmBtn.disabled = typed !== "DELETE";
  }
}

function ensureUniqueLabel(label, existing) {
  if (!label) return label;
  if (!existing || !existing.has(label)) return label;
  let idx = 2;
  let next = `${label} (${idx})`;
  while (existing.has(next)) {
    idx += 1;
    next = `${label} (${idx})`;
  }
  return next;
}

const SCIENCE_TERMS = [
  /cardio/i,
  /neuro/i,
  /pharma/i,
  /oncolog/i,
  /hemat/i,
  /genom/i,
  /immun/i,
  /metabol/i,
  /quantum/i,
  /biochem/i,
  /biome/i,
  /mri/i,
  /ct scan/i,
  /clinical/i,
  /diagnos/i,
  /therapy/i,
  /protocol/i,
  /pathway/i,
  /radiolog/i,
  /toxic/i,
  /antibody/i,
  /nanotech/i,
  /spectroscop/i,
  /algorithm/i,
  /statistic/i,
];
const SCIENCE_KEYWORDS = [
  "bio",
  "neuro",
  "cardio",
  "immun",
  "chem",
  "pharm",
  "gen",
  "thera",
  "clinic",
  "metabol",
  "protein",
  "cell",
  "quant",
  "astro",
  "nuclear",
  "optical",
  "scopic",
];

function applyInlineFormatting(text) {
  const frag = document.createDocumentFragment();
  const parts = text.split(/(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/);
  parts.forEach(part => {
    if (!part) return;
    if (part.startsWith("**") && part.endsWith("**") && part.length > 4) {
      const strong = document.createElement("strong");
      strong.textContent = part.slice(2, -2);
      frag.appendChild(strong);
      return;
    }
    if (part.startsWith("*") && part.endsWith("*") && part.length > 2) {
      const em = document.createElement("em");
      em.textContent = part.slice(1, -1);
      frag.appendChild(em);
      return;
    }
    if (part.startsWith("`") && part.endsWith("`") && part.length > 2) {
      const code = document.createElement("code");
      code.textContent = part.slice(1, -1);
      frag.appendChild(code);
      return;
    }
    frag.appendChild(document.createTextNode(part));
  });
  return frag;
}

function parseAssistantSections(rawText) {
  const expanded = expandDenseDelimiters(rawText);
  const lines = (expanded || "").split(/\r?\n/);
  const sections = [];
  let current = { heading: "Response", blocks: [] };
  let paragraph = null;

  const flushParagraph = () => {
    if (paragraph) {
      current.blocks.push({ type: "paragraph", text: paragraph.trim() });
      paragraph = null;
    }
  };

  lines.forEach(line => {
    const trimmed = line.trimEnd();
    const headingMatch = /^#{2,3}\s+(.+)/.exec(trimmed);
    if (headingMatch) {
      flushParagraph();
      if (current.blocks.length) {
        sections.push(current);
      }
      current = { heading: headingMatch[1].trim(), blocks: [] };
      return;
    }
    if (!trimmed) {
      flushParagraph();
      return;
    }
    const bulletMatch = /^[-•]\s+(.+)/.exec(trimmed);
    if (bulletMatch) {
      flushParagraph();
      const last = current.blocks[current.blocks.length - 1];
      if (last && last.type === "list") {
        last.items.push(bulletMatch[1]);
      } else {
        current.blocks.push({ type: "list", items: [bulletMatch[1]] });
      }
      return;
    }
    const orderedMatch = /^\d+[.)]\s+(.+)/.exec(trimmed);
    if (orderedMatch) {
      flushParagraph();
      const last = current.blocks[current.blocks.length - 1];
      if (last && last.type === "olist") {
        last.items.push(orderedMatch[1]);
      } else {
        current.blocks.push({ type: "olist", items: [orderedMatch[1]] });
      }
      return;
    }
    paragraph = paragraph ? `${paragraph} ${trimmed}` : trimmed;
  });

  flushParagraph();
  if (current.blocks.length) {
    sections.push(current);
  }
  return sections.length ? sections : [{ heading: "Response", blocks: [{ type: "paragraph", text: rawText }] }];
}

function expandDenseDelimiters(raw) {
  if (!raw) return raw;
  const pipeMatches = raw.match(/\s\|\s/g);
  const pipeCount = pipeMatches ? pipeMatches.length : 0;
  const looksLikeTable = /\n\s*\|[^|]+\|/.test(raw);
  if (pipeCount >= 2 && !looksLikeTable) {
    const segments = raw.split(/\s*\|\s*/).map(part => part.trim()).filter(Boolean);
    if (segments.length >= 3) {
      return segments
        .map(seg => (seg.startsWith("-") || seg.startsWith("•") ? seg : `- ${seg}`))
        .join("\n");
    }
  }
  return raw;
}

function renderAssistantStructured(text) {
  const wrapper = document.createElement("div");
  wrapper.className = "assistant-structured";
  const sections = parseAssistantSections(text);
  sections.forEach((section, idx) => {
    const block = document.createElement("section");
    block.className = "assistant-section";
    const heading = document.createElement("div");
    heading.className = "assistant-section-heading";
    heading.textContent = section.heading || `Section ${idx + 1}`;
    block.appendChild(heading);

    section.blocks.forEach(part => {
      if (part.type === "paragraph") {
        const p = document.createElement("p");
        p.className = "assistant-paragraph";
        p.appendChild(applyInlineFormatting(part.text));
        block.appendChild(p);
      } else if (part.type === "list") {
        const ul = document.createElement("ul");
        ul.className = "assistant-list";
        part.items.forEach(itemText => {
          const li = document.createElement("li");
          li.appendChild(applyInlineFormatting(itemText));
          ul.appendChild(li);
        });
        block.appendChild(ul);
      } else if (part.type === "olist") {
        const ol = document.createElement("ol");
        ol.className = "assistant-list ordered";
        part.items.forEach(itemText => {
          const li = document.createElement("li");
          li.appendChild(applyInlineFormatting(itemText));
          ol.appendChild(li);
        });
        block.appendChild(ol);
      }
    });

    wrapper.appendChild(block);
  });
  return wrapper;
}

function extractHashtags(text) {
  const found = new Set();
  const matches = text.match(/\b[\w-]{4,}\b/g) || [];
  SCIENCE_TERMS.forEach(re => {
    const match = text.match(re);
    if (match) {
      const cleaned = sanitizeTag(match[0]);
      if (cleaned) found.add(cleaned);
    }
  });
  matches.forEach(word => {
    const cleaned = sanitizeTag(word);
    if (!cleaned) return;
    if (isScientificWord(cleaned)) {
      found.add(cleaned);
    }
  });
  if (!found.size && matches.length) {
    matches
      .sort((a, b) => b.length - a.length)
      .slice(0, 2)
      .forEach(word => {
        const cleaned = sanitizeTag(word);
        if (cleaned) found.add(cleaned);
      });
  }
  return Array.from(found);
}

function sanitizeTag(value) {
  const normalized = value.replace(/[^a-z0-9]/gi, "").toLowerCase();
  if (!normalized) return "";
  return `#${normalized}`;
}

function isScientificWord(word) {
  const root = word.replace(/^#/, "");
  if (/\d/.test(root)) return true;
  return SCIENCE_KEYWORDS.some(keyword => root.includes(keyword));
}

const META_INLINE_STOPWORDS = new Set(["this", "that", "these", "those"]);
const META_STOPWORDS = new Set(["tell", "me", "different", "difference", "and", "best", "people"]);
const META_TOKEN_EDGE = /^[^a-z0-9]+|[^a-z0-9]+$/gi;

function stripMetaTokenEdges(value) {
  return value.replace(META_TOKEN_EDGE, "");
}

function normalizeMetaToken(value) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/['\u2019]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeMetaLookup(value) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/['\u2019]/g, " ")
    .replace(/[-\u2013\u2014]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeMetaLookupToken(value) {
  return normalizeMetaLookup(stripMetaTokenEdges(value));
}

function shouldJoinEntity(head, next) {
  const normalizedHead = normalizeMetaToken(head);
  const normalizedNext = normalizeMetaToken(next);
  if (!normalizedHead || !normalizedNext) return false;
  if (normalizedHead === "hepatitis" && /^[abc]$/.test(normalizedNext)) return true;
  if (normalizedHead === "vitamin" && /^[abcdk]$/.test(normalizedNext)) return true;
  if (normalizedHead === "type" && /^[0-9]$/.test(normalizedNext)) return true;
  if (normalizedHead === "stage" && /^[0-9]$/.test(normalizedNext)) return true;
  return false;
}

function extractMeaningfulTokens(raw) {
  if (typeof raw !== "string") return [];
  const trimmed = raw.trim();
  if (!trimmed) return [];
  const withoutMulti = trimmed.replace(MULTIWORD_PREPOSITION_REGEX, " ");
  const compact = withoutMulti.replace(/\s+/g, " ").trim();
  if (!compact) return [];
  const rawTokens = compact.split(" ");
  const joined = [];
  for (let i = 0; i < rawTokens.length; i += 1) {
    const current = stripMetaTokenEdges(rawTokens[i]);
    if (!current) continue;
    const nextRaw = rawTokens[i + 1];
    const next = nextRaw ? stripMetaTokenEdges(nextRaw) : "";
    if (next && shouldJoinEntity(current, next)) {
      joined.push(`${current} ${next}`);
      i += 1;
      continue;
    }
    joined.push(current);
  }
  const results = [];
  for (const token of joined) {
    const normalized = normalizeMetaToken(token);
    if (!normalized) continue;
    if (isArticleOrPreposition(normalized)) continue;
    if (META_INLINE_STOPWORDS.has(normalized) || META_STOPWORDS.has(normalized)) continue;
    results.push(token);
  }
  return results;
}

function cleanPromptPreserveOrder(raw) {
  if (typeof raw !== "string") return "";
  const trimmed = raw.trim();
  if (!trimmed) return "";
  const withoutMulti = trimmed.replace(MULTIWORD_PREPOSITION_REGEX, " ");
  const compact = withoutMulti.replace(/\s+/g, " ").trim();
  if (!compact) return "";
  const rawTokens = compact.split(" ");
  const kept = [];
  for (let i = 0; i < rawTokens.length; i += 1) {
    const token = rawTokens[i];
    if (!token) continue;
    let candidate = token;
    let skip = 0;
    const nextRaw = rawTokens[i + 1];
    if (nextRaw) {
      const currentJoin = stripMetaTokenEdges(token);
      const nextJoin = stripMetaTokenEdges(nextRaw);
      if (currentJoin && nextJoin && shouldJoinEntity(currentJoin, nextJoin)) {
        candidate = `${token} ${nextRaw}`;
        skip = 1;
      }
    }
    const lookup = normalizeMetaLookupToken(candidate);
    if (!lookup) {
      i += skip;
      continue;
    }
    if (isArticleOrPreposition(lookup)) {
      i += skip;
      continue;
    }
    if (META_INLINE_STOPWORDS.has(lookup) || META_STOPWORDS.has(lookup)) {
      i += skip;
      continue;
    }
    kept.push(candidate);
    i += skip;
  }
  return kept.join(" ").trim();
}

function isTrivialTopic(value) {
  const normalized = normalizeMetaToken(value).replace(/\s+/g, " ").trim();
  if (!normalized) return true;
  const compact = normalized.replace(/\s+/g, "");
  return compact.length <= 1;
}

function buildTopicPhrases(tokens, maxTopics = 5) {
  if (!Array.isArray(tokens) || !tokens.length || maxTopics <= 0) return [];
  const candidates = [];
  const addNgrams = (size) => {
    for (let i = 0; i <= tokens.length - size; i += 1) {
      const slice = tokens.slice(i, i + size);
      if (!slice.length) continue;
      candidates.push(slice.join(" "));
    }
  };
  addNgrams(3);
  addNgrams(2);
  addNgrams(1);
  const seen = new Set();
  const topics = [];
  for (const phrase of candidates) {
    if (isTrivialTopic(phrase)) continue;
    const key = normalizeMetaToken(phrase).replace(/\s+/g, " ").trim();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    topics.push(phrase);
    if (topics.length >= maxTopics) break;
  }
  return topics;
}

function cleanPromptToSentence(raw) {
  const tokens = extractMeaningfulTokens(raw);
  const cleaned = cleanPromptPreserveOrder(raw);
  return { cleaned, topics: buildTopicPhrases(tokens), tokens };
}

function ensureKaTeXWatcher() {
  if (katexIntervalStarted) return;
  katexIntervalStarted = true;
  const timer = setInterval(() => {
    if (window.renderMathInElement) {
      clearInterval(timer);
      while (pendingMathNodes.length) {
        const node = pendingMathNodes.shift();
        renderMathIfReady(node);
      }
    }
  }, 200);
}

function renderMathIfReady(node) {
  if (!node) return;
  if (window.renderMathInElement) {
    window.renderMathInElement(node, {
      delimiters: [
        { left: "\\[", right: "\\]", display: true },
        { left: "\\(", right: "\\)", display: false },
        { left: "$$", right: "$$", display: true },
        { left: "$", right: "$", display: false },
      ],
      throwOnError: false,
    });
  } else {
    pendingMathNodes.push(node);
    ensureKaTeXWatcher();
  }
}

function generateConversationId() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return `conv-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function loadStoredOcrSession() {
  const raw = safeGetItem(OCR_SESSION_KEY);
  if (!raw) return null;
  try {
    const data = JSON.parse(raw);
    if (data && typeof data === "object" && typeof data.fileHash === "string") {
      return data;
    }
  } catch {}
  return null;
}

function persistOcrSession(session) {
  if (!session) {
    safeSetItem(OCR_SESSION_KEY, "");
    return;
  }
  const payload = {
    fileHash: session.fileHash,
    extractedKey: session.extractedKey,
    pageCount: session.pageCount,
    processedRanges: session.processedRanges || [],
    filename: session.filename,
  };
  safeSetItem(OCR_SESSION_KEY, JSON.stringify(payload));
}

function clearOcrSession() {
  activePdfSession = null;
  safeSetItem(OCR_SESSION_KEY, "");
}

function loadStoredConversations() {
  // Primary: conversation index only
  const parseIndex = (raw) => {
    try {
      const data = JSON.parse(raw);
      if (!Array.isArray(data)) return [];
      return data
        .filter(entry => entry && typeof entry.id === "string")
        .map(entry => ({
          id: entry.id,
          title: entry.title || "Conversation",
          createdAt: typeof entry.createdAt === "number" ? entry.createdAt : Date.now(),
          updatedAt: typeof entry.updatedAt === "number" ? entry.updatedAt : Date.now(),
          selectedDocId: entry.selectedDocId || null,
          selectedDocTitle: entry.selectedDocTitle || "",
        }));
    } catch {
      return [];
    }
  };

  const idxRaw = safeGetItem(CONVO_INDEX_KEY);
  if (idxRaw) {
    return parseIndex(idxRaw);
  }

  // Migration: legacy single blob storage -> index + per-conversation entries
  const legacy = safeGetItem(CONVO_STORAGE_KEY);
  if (!legacy) return [];
  let migratedIndex = [];
  try {
    const data = JSON.parse(legacy);
    if (Array.isArray(data)) {
      migratedIndex = data
        .filter(conv => conv && typeof conv.id === "string")
        .map(conv => {
          const sanitized = shrinkConversationForStorage(conv);
          if (sanitized) {
            saveConversationDocument(sanitized);
          }
          return {
            id: conv.id,
            title: conv.title || "Conversation",
            createdAt: conv.createdAt || Date.now(),
            updatedAt: conv.updatedAt || Date.now(),
            selectedDocId: conv.selectedDocId || null,
            selectedDocTitle: conv.selectedDocTitle || "",
          };
        })
        .filter(Boolean);
      saveStoredConversations(migratedIndex);
      safeRemoveItem(CONVO_STORAGE_KEY);
    }
  } catch {
    migratedIndex = [];
  }
  return migratedIndex;
}

function loadTitleOverrides() {
  const raw = safeGetItem(TITLE_OVERRIDE_KEY);
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function saveTitleOverrides(overrides) {
  if (!overrides || typeof overrides !== "object") return;
  safeSetItem(TITLE_OVERRIDE_KEY, JSON.stringify(overrides));
}

function loadActiveConversationId() {
  const raw = safeGetItem(LAST_CONVO_KEY) || safeGetItem(LEGACY_LAST_CONVO_KEY);
  return typeof raw === "string" && raw.trim() ? raw.trim() : "";
}

function saveActiveConversationId(id) {
  if (!id) return;
  safeSetItem(LAST_CONVO_KEY, id);
  safeSetItem(LEGACY_LAST_CONVO_KEY, id);
}

function safeGetItem(key) {
  try {
    return localStorage.getItem(key);
  } catch {
  }
  try {
    return sessionStorage.getItem(key);
  } catch {
  }
  return null;
}

function safeSetItem(key, value) {
  try {
    localStorage.setItem(key, value);
    return true;
  } catch {
  }
  try {
    sessionStorage.setItem(key, value);
    return true;
  } catch {
  }
  return false;
}

function safeRemoveItem(key) {
  try {
    localStorage.removeItem(key);
    return true;
  } catch {}
  try {
    sessionStorage.removeItem(key);
    return true;
  } catch {}
  return false;
}

function saveStoredConversations(index = storedConversations) {
  const payload = (index || []).slice(0, 200).map(entry => ({
    id: entry.id,
    title: entry.title || "Conversation",
    createdAt: entry.createdAt || Date.now(),
    updatedAt: entry.updatedAt || Date.now(),
    selectedDocId: entry.selectedDocId || null,
    selectedDocTitle: entry.selectedDocTitle || "",
  }));
  if (safeSetItem(CONVO_INDEX_KEY, JSON.stringify(payload))) return;
  safeSetItem(CONVO_INDEX_KEY, JSON.stringify(payload.map(entry => ({ id: entry.id, title: entry.title, updatedAt: entry.updatedAt }))));
}

function conversationStorageKey(id) {
  return `${CONVO_PREFIX}${id}`;
}

function loadConversationDocument(id) {
  if (!id) return null;
  const raw = safeGetItem(conversationStorageKey(id));
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || parsed.id !== id) return null;
    const safe = shrinkConversationForStorage(parsed);
    if (DEBUG_PERSIST && safe) {
      console.log("[PERSIST_LOAD]", { conv: id, messages: (safe.messages || []).length });
    }
    return safe || null;
  } catch {
    return null;
  }
}

function saveConversationDocument(conv) {
  const safe = shrinkConversationForStorage(conv);
  if (!safe) return;
  const payload = JSON.stringify(safe);
  safeSetItem(conversationStorageKey(conv.id), payload);
  if (DEBUG_PERSIST) {
    console.log("[PERSIST_SAVE]", {
      conv: conv.id,
      messages: safe.messages.length,
      bytes: payload.length,
    });
  }
}

function deleteConversationDocument(id) {
  if (!id) return;
  safeRemoveItem(conversationStorageKey(id));
}

function isPlainObject(value) {
  return value && typeof value === "object" && !Array.isArray(value);
}

function normalizeConversationIndexEntry(raw) {
  if (!raw || typeof raw !== "object") return null;
  const id = typeof raw.id === "string" ? raw.id.trim() : "";
  if (!id) return null;
  const createdAt = typeof raw.createdAt === "number" ? raw.createdAt : Date.now();
  const updatedAt = typeof raw.updatedAt === "number" ? raw.updatedAt : createdAt;
  return {
    id,
    title: typeof raw.title === "string" ? raw.title : "Conversation",
    createdAt,
    updatedAt,
    selectedDocId: typeof raw.selectedDocId === "string" ? raw.selectedDocId : null,
    selectedDocTitle: typeof raw.selectedDocTitle === "string" ? raw.selectedDocTitle : "",
  };
}

function mergeConversationIndexes(localIndex, remoteIndex) {
  const map = new Map();
  (localIndex || []).forEach(entry => {
    if (entry?.id) map.set(entry.id, entry);
  });
  (remoteIndex || []).forEach(entry => {
    if (!entry?.id) return;
    const existing = map.get(entry.id);
    if (!existing) {
      map.set(entry.id, entry);
      return;
    }
    const updatedAt = Math.max(existing.updatedAt || 0, entry.updatedAt || 0);
    const createdAt = Math.min(existing.createdAt || updatedAt, entry.createdAt || updatedAt);
    const winner = (entry.updatedAt || 0) >= (existing.updatedAt || 0) ? entry : existing;
    map.set(entry.id, {
      ...existing,
      ...entry,
      title: winner.title || existing.title || entry.title || "Conversation",
      createdAt,
      updatedAt,
      selectedDocId: winner.selectedDocId ?? existing.selectedDocId ?? entry.selectedDocId ?? null,
      selectedDocTitle: winner.selectedDocTitle ?? existing.selectedDocTitle ?? entry.selectedDocTitle ?? "",
    });
  });
  return Array.from(map.values()).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
}

function normalizeConversationMessageForDisplay(raw) {
  if (!raw || typeof raw !== "object") return null;
  const meta = isPlainObject(raw.metadata) ? raw.metadata : {};
  const merged = { ...meta, ...raw };
  const role = merged.role === "assistant" ? "assistant" : merged.role === "system" ? "system" : "user";
  const text = typeof merged.text === "string"
    ? merged.text
    : typeof merged.content === "string"
      ? merged.content
      : "";
  const msg = {
    id: typeof merged.id === "string" ? merged.id : generateMessageId(),
    role,
    text,
    createdAt: typeof merged.createdAt === "number" ? merged.createdAt : Date.now(),
  };
  if (typeof merged.at === "string") msg.at = merged.at;
  if (typeof merged.model === "string") msg.model = merged.model;
  if (Array.isArray(merged.attachments)) {
    msg.attachments = cloneMessageAttachments(merged.attachments);
  }
  if (typeof merged.imageUrl === "string") msg.imageUrl = merged.imageUrl;
  if (typeof merged.imageAlt === "string") msg.imageAlt = merged.imageAlt;
  if (merged.docId) msg.docId = merged.docId;
  if (merged.docTitle) msg.docTitle = merged.docTitle;
  if (merged.lectureId) msg.lectureId = merged.lectureId;
  if (merged.extractedKey) msg.extractedKey = merged.extractedKey;
  if (merged.requestId) msg.requestId = merged.requestId;
  if (Array.isArray(merged.references)) msg.references = merged.references.slice(0, 50);
  if (Array.isArray(merged.evidence)) msg.evidence = merged.evidence.slice(0, 50);
  if (typeof merged.renderedMarkdown === "string") msg.renderedMarkdown = merged.renderedMarkdown;
  if (Array.isArray(merged.sources)) {
    const normalized = normalizeStoredSources(merged.sources);
    if (normalized.length) msg.sources = normalized;
  }
  if (Array.isArray(merged.citations)) {
    const normalized = normalizeStoredCitations(merged.citations);
    if (normalized.length) msg.citations = normalized;
  }
  if (Array.isArray(merged.answerSegments)) {
    const sanitized = sanitizeAnswerSegmentsForStorage(merged.answerSegments);
    if (sanitized.length) msg.answerSegments = sanitized;
  }
  if (typeof merged.rawPrompt === "string") msg.rawPrompt = merged.rawPrompt;
  if (typeof merged.cleanedPrompt === "string") msg.cleanedPrompt = merged.cleanedPrompt;
  if (Array.isArray(merged.topics)) msg.topics = merged.topics.slice(0, 20);
  return msg;
}

function normalizeConversationDocument(raw) {
  if (!raw || typeof raw !== "object") return null;
  const id = typeof raw.id === "string" ? raw.id : "";
  if (!id) return null;
  const messages = Array.isArray(raw.messages)
    ? raw.messages.map(normalizeConversationMessageForDisplay).filter(Boolean)
    : [];
  return {
    id,
    title: typeof raw.title === "string" ? raw.title : "Conversation",
    createdAt: typeof raw.createdAt === "number" ? raw.createdAt : Date.now(),
    updatedAt: typeof raw.updatedAt === "number" ? raw.updatedAt : Date.now(),
    selectedDocId: typeof raw.selectedDocId === "string" ? raw.selectedDocId : null,
    selectedDocTitle: typeof raw.selectedDocTitle === "string" ? raw.selectedDocTitle : "",
    truncated: Boolean(raw.truncated),
    messages,
  };
}

function mergeConversationDocuments(localDoc, remoteDoc) {
  if (!remoteDoc) return localDoc;
  if (!localDoc) return remoteDoc;
  const localMessages = Array.isArray(localDoc.messages) ? localDoc.messages : [];
  const remoteMessages = Array.isArray(remoteDoc.messages) ? remoteDoc.messages : [];
  const base = remoteMessages.length >= localMessages.length ? remoteMessages : localMessages;
  const extras = base === remoteMessages ? localMessages : remoteMessages;
  const merged = base.map(msg => ({ ...msg }));
  const indexById = new Map();
  merged.forEach((msg, idx) => {
    if (msg?.id) indexById.set(msg.id, idx);
  });
  extras.forEach(msg => {
    if (!msg?.id) return;
    if (indexById.has(msg.id)) {
      merged[indexById.get(msg.id)] = msg;
      return;
    }
    indexById.set(msg.id, merged.length);
    merged.push(msg);
  });
  const updatedAt = Math.max(localDoc.updatedAt || 0, remoteDoc.updatedAt || 0);
  const createdAt = Math.min(localDoc.createdAt || updatedAt, remoteDoc.createdAt || updatedAt);
  const winner = (remoteDoc.updatedAt || 0) >= (localDoc.updatedAt || 0) ? remoteDoc : localDoc;
  return {
    id: localDoc.id || remoteDoc.id,
    title: winner.title || localDoc.title || remoteDoc.title || "Conversation",
    createdAt,
    updatedAt,
    selectedDocId: winner.selectedDocId ?? localDoc.selectedDocId ?? remoteDoc.selectedDocId ?? null,
    selectedDocTitle: winner.selectedDocTitle ?? localDoc.selectedDocTitle ?? remoteDoc.selectedDocTitle ?? "",
    truncated: Boolean(localDoc.truncated || remoteDoc.truncated),
    messages: dedupeMessages(merged),
  };
}

function serializeConversationMessageForRemote(msg) {
  const safe = sanitizeMessageForStorage(msg);
  if (!safe) return null;
  const metadata = {};
  if (typeof safe.model === "string") metadata.model = safe.model;
  if (Array.isArray(safe.attachments)) metadata.attachments = cloneMessageAttachments(safe.attachments);
  if (typeof safe.imageUrl === "string") metadata.imageUrl = safe.imageUrl;
  if (typeof safe.imageAlt === "string") metadata.imageAlt = safe.imageAlt;
  if (safe.docId) metadata.docId = safe.docId;
  if (safe.docTitle) metadata.docTitle = safe.docTitle;
  if (safe.lectureId) metadata.lectureId = safe.lectureId;
  if (safe.extractedKey) metadata.extractedKey = safe.extractedKey;
  if (safe.requestId) metadata.requestId = safe.requestId;
  if (Array.isArray(safe.references)) metadata.references = safe.references.slice(0, 50);
  if (Array.isArray(safe.evidence)) metadata.evidence = safe.evidence.slice(0, 50);
  if (typeof safe.renderedMarkdown === "string") metadata.renderedMarkdown = safe.renderedMarkdown;
  if (Array.isArray(safe.sources)) metadata.sources = normalizeStoredSources(safe.sources);
  if (Array.isArray(safe.citations)) metadata.citations = normalizeStoredCitations(safe.citations);
  if (Array.isArray(safe.answerSegments)) metadata.answerSegments = sanitizeAnswerSegmentsForStorage(safe.answerSegments);
  if (typeof safe.rawPrompt === "string") metadata.rawPrompt = safe.rawPrompt;
  if (typeof safe.cleanedPrompt === "string") metadata.cleanedPrompt = safe.cleanedPrompt;
  if (Array.isArray(safe.topics)) metadata.topics = safe.topics.slice(0, 20);
  return {
    id: safe.id,
    role: safe.role,
    content: typeof safe.text === "string" ? safe.text : "",
    createdAt: typeof safe.createdAt === "number" ? safe.createdAt : Date.now(),
    metadata: Object.keys(metadata).length ? metadata : undefined,
  };
}

function serializeConversationForRemote(conv) {
  if (!conv || !conv.id) return null;
  const safe = shrinkConversationForStorage(conv);
  if (!safe) return null;
  const messages = (safe.messages || [])
    .map(serializeConversationMessageForRemote)
    .filter(Boolean);
  return {
    id: safe.id,
    title: safe.title || "Conversation",
    createdAt: safe.createdAt || Date.now(),
    updatedAt: safe.updatedAt || Date.now(),
    selectedDocId: safe.selectedDocId || null,
    selectedDocTitle: safe.selectedDocTitle || "",
    truncated: Boolean(safe.truncated),
    messages,
  };
}

function scheduleConversationSync(entry) {
  if (!entry || !entry.id) return;
  pendingConversationSync.set(entry.id, entry);
  if (conversationSyncTimer) return;
  conversationSyncTimer = setTimeout(() => {
    conversationSyncTimer = null;
    const batch = Array.from(pendingConversationSync.values());
    pendingConversationSync.clear();
    batch.forEach(item => syncConversationToServer(item));
  }, CONVO_REMOTE_SYNC_DEBOUNCE_MS);
}

async function syncConversationToServer(entry) {
  const payload = serializeConversationForRemote(entry);
  if (!payload || !Array.isArray(payload.messages) || !payload.messages.length) return;
  try {
    const res = await fetch(CONVO_REMOTE_ENDPOINT, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ conversation: payload }),
    });
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      console.warn("[CONVO_SYNC_FAILED]", res.status, text);
      return;
    }
    const data = await res.json().catch(() => ({}));
    const remoteDoc = normalizeConversationDocument(data?.conversation);
    if (remoteDoc) {
      saveConversationDocument(remoteDoc);
    }
    const indexEntry = normalizeConversationIndexEntry(data?.index || {});
    if (indexEntry) {
      storedConversations = mergeConversationIndexes(storedConversations, [indexEntry]);
      saveStoredConversations();
      renderConversationList();
    }
  } catch (err) {
    console.warn("[CONVO_SYNC_ERROR]", err);
  }
}

async function syncConversationIndexFromServer(opts = {}) {
  const force = Boolean(opts.force);
  if (!force && Date.now() - lastRemoteIndexFetch < CONVO_REMOTE_INDEX_TTL_MS) return;
  lastRemoteIndexFetch = Date.now();
  try {
    const res = await fetch(CONVO_REMOTE_ENDPOINT, {
      method: "GET",
      headers: { "accept": "application/json" },
      cache: "no-store",
    });
    if (!res.ok) return;
    const payload = await res.json().catch(() => ({}));
    const remote = Array.isArray(payload?.conversations)
      ? payload.conversations.map(normalizeConversationIndexEntry).filter(Boolean)
      : [];
    if (!remote.length) return;
    storedConversations = mergeConversationIndexes(storedConversations, remote);
    saveStoredConversations();
    renderConversationList();
    if (opts.bootstrap && !loadActiveConversationId() && storedConversations.length) {
      loadConversation(storedConversations[0].id);
    }
  } catch (err) {
    console.warn("[CONVO_INDEX_SYNC_FAILED]", err);
  }
}

function backfillLocalConversationsToServer(limit = 12) {
  (storedConversations || []).slice(0, limit).forEach(entry => {
    if (!entry?.id) return;
    const doc = loadConversationDocument(entry.id);
    if (doc && Array.isArray(doc.messages) && doc.messages.length) {
      scheduleConversationSync(doc);
    }
  });
}

async function fetchConversationDocumentFromServer(id) {
  if (!id) return null;
  try {
    const res = await fetch(`${CONVO_REMOTE_ENDPOINT}/${encodeURIComponent(id)}`, {
      method: "GET",
      headers: { "accept": "application/json" },
      cache: "no-store",
    });
    if (!res.ok) return null;
    const payload = await res.json().catch(() => ({}));
    return normalizeConversationDocument(payload?.conversation || payload);
  } catch (err) {
    console.warn("[CONVO_FETCH_FAILED]", err);
    return null;
  }
}

async function deleteConversationRemote(id) {
  if (!id) return;
  try {
    await fetch(`${CONVO_REMOTE_ENDPOINT}/${encodeURIComponent(id)}`, {
      method: "DELETE",
    });
  } catch (err) {
    console.warn("[CONVO_DELETE_FAILED]", err);
  }
}

function showConversationWarning(id, message) {
  if (!chatLog || warnedConversations.has(id)) return;
  warnedConversations.add(id);
  const notice = document.createElement("div");
  notice.className = "status-pill";
  notice.dataset.state = "warning";
  notice.textContent = message;
  chatLog.prepend(notice);
  try {
    const raw = safeGetItem(CONVO_REMOTE_WARN_KEY) || "[]";
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && !parsed.includes(id)) {
      parsed.push(id);
      safeSetItem(CONVO_REMOTE_WARN_KEY, JSON.stringify(parsed.slice(-100)));
    }
  } catch {}
}

function collectCitations(evt, store) {
  const ingest = (payload) => {
    normalizeCitationsInput(payload).forEach(item => store.push(item));
  };
  const ingestAnnotations = (annotations) => {
    if (!Array.isArray(annotations)) return;
    annotations.forEach(ann => {
      ingest(ann);
      ingest(ann?.citations);
    });
  };

  [
    evt.citations,
    evt.delta?.citations,
    evt.response?.citations,
  ].forEach(ingest);

  if (Array.isArray(evt.choices)) {
    evt.choices.forEach(choice => ingest(choice?.delta?.citations));
  }

  if (Array.isArray(evt.response?.output)) {
    evt.response.output.forEach(block => {
      block.content?.forEach(part => ingest(part?.citations));
      block.content?.forEach(part => {
        ingest(part?.citation_metadata?.citations || part?.citation_metadata);
        ingestAnnotations(part?.annotations);
      });
    });
  }
}

function parseCitationLines(text) {
  const entries = [];
  const regex = /(?:^|\n)\s*[#\[]?\s*(\d+)[\]\).:]?\s*[-–—:]?\s*(https?:\/\/\S+)/gi;
  let match;
  while ((match = regex.exec(text))) {
    const id = match[1];
    const url = normalizeCitationUrl(match[2]);
    if (url) entries.push({ id, url });
  }
  return entries;
}

function normalizeCitationsInput(raw) {
  const normalized = [];
  const isCitationShape = (obj) => {
    if (!obj || typeof obj !== "object") return false;
    return Boolean(
      obj.url ||
      obj.href ||
      obj.link ||
      obj.source ||
      obj.source_url ||
      obj.file_id ||
      obj.file_citation ||
      obj.file_path ||
      obj.citation_id ||
      obj.id ||
      obj.i
    );
  };

  const visit = (entry, keyHint, indexHint) => {
    if (!entry) return;
    if (typeof entry === "string") {
      const trimmed = entry.trim();
      if (!trimmed) return;
      if (/^[\[{]/.test(trimmed)) {
        try {
          visit(JSON.parse(trimmed), keyHint, indexHint);
          return;
        } catch {
          // fall through
        }
      }
      const linePairs = parseCitationLines(trimmed);
      if (linePairs.length) {
        linePairs.forEach(pair => visit(pair, keyHint, indexHint));
        return;
      }
      const url = normalizeCitationUrl(trimmed);
      if (!url) return;
      normalized.push({ url, id: keyHint || String(indexHint ?? normalized.length + 1) });
      return;
    }
    if (typeof entry === "number") {
      const url = normalizeCitationUrl(String(entry));
      if (url) normalized.push({ url, id: keyHint || String(indexHint ?? normalized.length + 1) });
      return;
    }
    if (Array.isArray(entry)) {
      entry.forEach((child, idx) => visit(child, keyHint, idx + 1));
      return;
    }
    if (typeof entry === "object") {
      if (isCitationShape(entry)) {
        const candidate = { ...entry };
        const fallbackId = keyHint || (typeof indexHint === "number" ? String(indexHint) : undefined);
        if (!candidate.id && fallbackId) candidate.id = fallbackId;
        if (!candidate.i && fallbackId && /^\d+$/.test(fallbackId)) {
          candidate.i = Number(fallbackId);
        }
        normalized.push(candidate);
        return;
      }
      const pairs = Object.entries(entry);
      if (!pairs.length) return;
      pairs.forEach(([k, v], idx) => {
        if (typeof v === "object") {
          visit(v, k, idx + 1);
        } else {
          visit({ id: k, url: v }, k, idx + 1);
        }
      });
    }
  };

  visit(raw);
  return normalized;
}

function normalizeCitationUrl(raw) {
  if (typeof raw !== "string") return "";
  const trimmed = raw.trim();
  if (!trimmed || trimmed === "#") return "";
  const stripped = trimmed.replace(/^[\[\(]+/, "").replace(/[)\]\.,;!?]+$/, "");
  if (/^[a-z][a-z0-9+\-.]*:\/\//i.test(stripped)) return stripped;
  if (stripped.startsWith("www.")) return `https://${stripped}`;
  return stripped;
}

function buildCitationPlaceholder(id) {
  return `${CITATION_PLACEHOLDER_PREFIX}${id}__`;
}

function buildSourcesSection(sources) {
  if (!Array.isArray(sources) || !sources.length) return null;
  const section = document.createElement("div");
  section.className = "assistant-section sources-section";
  const heading = document.createElement("div");
  heading.className = "section-heading";
  heading.textContent = `${pickEmoji("reference")} Sources`;
  const body = document.createElement("div");
  body.className = "section-body";
  const list = document.createElement("ol");
  list.className = "sources-list";
  let hasAny = false;
  sources.forEach(src => {
    if (!src || !src.url || !isLikelyValidCitationUrl(src.url)) return;
    const li = document.createElement("li");
    li.value = src.id;
    const link = document.createElement("a");
    link.className = "sources-link";
    link.href = src.url;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    const label = src.title ? src.title : (src.domain || src.url);
    link.textContent = label;
    link.setAttribute("aria-label", `Source ${src.id}: ${label}`);
    link.title = src.url;
    li.appendChild(link);
    if (src.title && src.domain) {
      const domain = document.createElement("span");
      domain.className = "sources-domain";
      domain.textContent = ` — ${src.domain}`;
      li.appendChild(domain);
    }
    list.appendChild(li);
    hasAny = true;
  });
  if (!hasAny) return null;
  body.appendChild(list);
  section.append(heading, body);
  return section;
}

function renderCitedAnswerSegments(bubble, segments, sources) {
  if (!bubble) return;
  const citationMap = buildCitationMapFromSegments(segments);
  const textParts = [];
  (segments || []).forEach(seg => {
    if (!seg) return;
    if (seg.type === "text") {
      textParts.push(typeof seg.text === "string" ? seg.text : "");
    } else if (seg.type === "citation") {
      const id = Number(seg.id);
      if (Number.isFinite(id)) {
        textParts.push(buildCitationPlaceholder(id));
      }
    }
  });
  const mergedText = textParts.join("");
  const textBlock = ensureBubbleTextNode(bubble, { reset: true });
  const plainAnswer = buildCitedPlainText(segments);
  const msgId = bubble?.dataset?.msgId;
  const normalizedSources = normalizeCitedSources(sources, citationMap);
  const container = decorateAssistantResponse(mergedText, citationMap, {
    sources: normalizedSources,
    answerText: plainAnswer,
    msgId,
    preserveCitationMap: true,
  });
  replaceCitationPlaceholders(container, citationMap, CITATION_PLACEHOLDER_PREFIX);
  textBlock.replaceChildren(container);
  enhanceAutoTables(textBlock);
  enhanceTablesIn(textBlock);
  finalizeResponseCard(container);
  bubble.classList.toggle("has-response", true);
  renderMathIfReady(bubble);
}

function buildCitedFallbackText(segments, sources) {
  const parts = [];
  (segments || []).forEach(seg => {
    if (!seg) return;
    if (seg.type === "text") {
      parts.push(typeof seg.text === "string" ? seg.text : "");
    } else if (seg.type === "citation") {
      const id = Number(seg.id);
      parts.push(Number.isFinite(id) ? `[${id}]` : "");
    }
  });
  let text = parts.join("");
  const fallbackSources = normalizeCitedSources(sources, buildCitationMapFromSegments(segments));
  if (fallbackSources.length) {
    const lines = fallbackSources.map(src => {
      const label = src.title ? `${src.title} — ${src.url}` : src.url;
      return `${src.id}. ${label}`;
    });
    text = `${text}\n\nSources:\n${lines.join("\n")}`;
  }
  return text.trim();
}

function buildCitedPlainText(segments) {
  const parts = [];
  (segments || []).forEach(seg => {
    if (!seg) return;
    if (seg.type === "text") {
      parts.push(typeof seg.text === "string" ? seg.text : "");
    } else if (seg.type === "citation") {
      const id = Number(seg.id);
      parts.push(Number.isFinite(id) ? `[${id}]` : "");
    }
  });
  return parts.join("").trim();
}

function extractFileId(cite) {
  if (!cite || typeof cite !== "object") return "";
  if (typeof cite.file_id === "string" && cite.file_id) return cite.file_id;
  if (typeof cite.fileId === "string" && cite.fileId) return cite.fileId;
  if (cite.file_path?.file_id) return cite.file_path.file_id;
  if (cite.file_citation?.file_id) return cite.file_citation.file_id;
  if (cite.file_reference?.file_id) return cite.file_reference.file_id;
  return "";
}

function buildCitationMap(citations) {
  const map = new Map();
  normalizeCitationsInput(citations).forEach((cite, idx) => {
    if (!cite) return;
    const id = String(
      cite.id ??
      cite.i ??
      cite.index ??
      cite.number ??
      cite.citation_id ??
      idx + 1,
    );
    const title = typeof cite.title === "string" ? cite.title : "";
    const snippet = typeof cite.snippet === "string" ? cite.snippet : "";
    const domain = typeof cite.domain === "string" ? cite.domain : "";
    const urlCandidate =
      cite.url ||
      cite.href ||
      cite.link ||
      cite.source ||
      cite.source_url ||
      (extractFileId(cite) && attachmentUrlMap[extractFileId(cite)]) ||
      null;
    const url = normalizeCitationUrl(urlCandidate);
    if (isLikelyValidCitationUrl(url)) {
      map.set(id, { url, title, snippet, domain });
    }
  });
  return map;
}

function serializeCitationMap(citationMap) {
  const entries = [];
  if (!(citationMap instanceof Map)) return entries;
  citationMap.forEach((value, key) => {
    const idValue = value?.id ?? key;
    const id = Number(idValue);
    const url = typeof value === "string" ? value : value?.url;
    if (!Number.isFinite(id) || !url) return;
    entries.push({
      id,
      url,
      title: typeof value?.title === "string" ? value.title : "",
      snippet: typeof value?.snippet === "string" ? value.snippet : "",
      domain: typeof value?.domain === "string" ? value.domain : "",
    });
  });
  entries.sort((a, b) => a.id - b.id);
  return entries.slice(0, MAX_STORED_CITATIONS);
}

/**
 * Build a citation map from stored citation/sources metadata.
 *
 * @param params - Stored citations/sources arrays.
 * @returns Map of citation id to URL metadata.
 */
function buildCitationMapFromStored({ citations, sources } = {}) {
  if (Array.isArray(citations) && citations.length) {
    return buildCitationMap(citations);
  }
  if (Array.isArray(sources) && sources.length) {
    return buildCitationMap(sources);
  }
  return new Map();
}

function enrichCitationMapFromText(text, map) {
  const refRegex = /(\d+)\.\s.*?(https?:\/\/\S+)/g;
  let match;
  while ((match = refRegex.exec(text))) {
    const id = match[1];
    let url = match[2];
    url = url.replace(/[\]\)]?[\.,;!?]*$/, "");
    if (isLikelyValidCitationUrl(url)) {
      const existing = map.get(id);
      if (existing && typeof existing === "object") {
        map.set(id, { ...existing, url });
      } else {
        map.set(id, url);
      }
    }
  }
  parseCitationLines(text).forEach(({ id, url }) => {
    if (id && isLikelyValidCitationUrl(url)) {
      const key = String(id);
      const existing = map.get(key);
      if (existing && typeof existing === "object") {
        map.set(key, { ...existing, url });
      } else {
        map.set(key, url);
      }
    }
  });
}

function downloadExport(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function exportConversation(id, format = "json") {
  const entry = storedConversations.find(conv => conv.id === id) || {};
  const doc = loadConversationDocument(id) || { id, messages: [] };
  const payload = {
    id,
    title: getConversationTitle({ id, title: entry.title || doc.title || "Conversation" }),
    createdAt: doc.createdAt || entry.createdAt || Date.now(),
    updatedAt: doc.updatedAt || entry.updatedAt || Date.now(),
    messages: doc.messages || [],
  };
  if (format === "markdown") {
    const lines = [
      `# ${payload.title}`,
      "",
      `- Created: ${new Date(payload.createdAt).toLocaleString()}`,
      `- Updated: ${new Date(payload.updatedAt).toLocaleString()}`,
      "",
      "## Transcript",
      "",
    ];
    payload.messages.forEach(msg => {
      const stamp = msg.createdAt ? new Date(msg.createdAt).toLocaleString() : "";
      const role = msg.role === "assistant" ? "Assistant" : "User";
      lines.push(`### ${role}${stamp ? ` (${stamp})` : ""}`);
      lines.push(msg.text || msg.content || "");
      lines.push("");
    });
    downloadExport(`${payload.title.replace(/\\s+/g, "_")}.md`, lines.join("\\n"), "text/markdown");
    return;
  }
  downloadExport(`${payload.title.replace(/\\s+/g, "_")}.json`, JSON.stringify(payload, null, 2), "application/json");
}

function deleteConversation(id) {
  if (!id) return;
  deleteConversationDocument(id);
  storedConversations = storedConversations.filter(conv => conv.id !== id);
  setConversationTitleOverride(id, "", { skipSync: true, skipPersist: true });
  saveStoredConversations();
  deleteConversationRemote(id);
  if (activeConversationId === id) {
    if (storedConversations.length) {
      loadConversation(storedConversations[0].id);
    } else {
      startNewConversation();
    }
  }
  renderConversationList();
}

function backfillCitationMapFromUrls(text, map) {
  const idsInText = Array.from(new Set(Array.from(text.matchAll(/\[(\d+)\]/g)).map(m => m[1])));
  if (!idsInText.length) return;
  const urlRegex = /https?:\/\/[^\s<>\]\)]*/g;
  const urls = [];
  let match;
  while ((match = urlRegex.exec(text))) {
    const url = normalizeCitationUrl(match[0]);
    if (isLikelyValidCitationUrl(url) && !urls.includes(url)) urls.push(url);
  }
  idsInText.forEach((id, idx) => {
    if (map.has(id)) return;
    const fallbackUrl = urls[idx] || urls[urls.length - 1];
    if (isLikelyValidCitationUrl(fallbackUrl)) {
      map.set(id, fallbackUrl);
    }
  });
}

function collectUrlsFromText(text) {
  const urls = [];
  const urlRegex = /https?:\/\/[^\s<>\]\)]*/gi;
  let match;
  while ((match = urlRegex.exec(text || ""))) {
    const url = normalizeCitationUrl(match[0]);
    if (isLikelyValidCitationUrl(url) && !urls.includes(url)) urls.push(url);
  }
  return urls;
}

function stripTrailingCitationNoise(text) {
  const markers = [
    /:::\s*citations?\b/i,
    /::\s*citations?\b/i,
    /\[\s*\{\s*"i"\s*:/i,
    /\{\s*"i"\s*:\s*\d+/i,
  ];
  let cut = text.length;
  markers.forEach((re) => {
    const m = re.exec(text || "");
    if (m && m.index < cut) cut = m.index;
  });
  if (cut < text.length) {
    return { cleanText: text.slice(0, cut).trim(), citations: [] };
  }
  return null;
}

function extractCitationBlock(text) {
  const citeBlockRegex = /:::\s*citations\s*([\s\S]*?)\s*:::/i;
  const match = citeBlockRegex.exec(text || "");
  if (!match) {
    const stripped = stripTrailingCitationNoise(text);
    if (stripped) return stripped;
    return { cleanText: text, citations: [] };
  }

  const rawBlock = match[1] || "";
  const cleanedText = [
    text.slice(0, match.index).trim(),
    text.slice(match.index + match[0].length).trim(),
  ].filter(Boolean).join("\n\n") || text;

  const tryParse = (payload) => {
    try {
      const parsed = JSON.parse(payload);
      return normalizeCitationsInput(parsed);
    } catch {
      return [];
    }
  };

  let parsed = tryParse(rawBlock.trim());
  if (!parsed.length) {
    const bracketMatch = rawBlock.match(/\[[\s\S]*\]/);
    if (bracketMatch?.[0]) {
      const fallback = tryParse(bracketMatch[0]);
      if (fallback.length) {
        parsed = fallback;
      }
    }
  }
  if (!parsed.length) {
    parsed = parseCitationLines(rawBlock);
  }

  return { cleanText: cleanedText, citations: parsed };
}

function normalizeAssistantText(text) {
  if (!text) return "";
  let normalized = text.trim();
  if (normalized.includes("\\n") && !normalized.includes("\n")) {
    const unescaped = normalized.replace(/\\n/g, "\n");
    normalized = unescaped.trim();
  }
  return normalized;
}

function isSourcesUnavailableBlock(block) {
  const trimmed = (block || "").trim();
  if (!trimmed) return false;
  return trimmed === SOURCES_UNAVAILABLE_NOTE;
}

function isReferenceBlock(block) {
  const trimmed = (block || "").trim();
  if (!trimmed) return false;
  const lower = trimmed.toLowerCase();
  if (/^references?\b/.test(lower)) return true;
  const lines = trimmed.split(/\n/).filter(Boolean);
  const urlLines = lines.filter(line => /https?:\/\//i.test(line));
  const numberedLines = lines.filter(line => /^\s*\d+[\).\s]/.test(line));
  if (urlLines.length >= 2 && urlLines.length >= lines.length * 0.5) return true;
  if (numberedLines.length >= 2 && urlLines.length) return true;
  return false;
}

function parseInlineReferenceItems(block) {
  const cleaned = (block || "").replace(/\r/g, "").trim();
  if (!cleaned) return [];
  const lines = cleaned.split(/\n/).map(line => line.trim()).filter(Boolean);
  const items = [];
  const inlinePattern = /(\d+)\s*(?:[.)]|-)\s+/g;

  const pushItem = (value) => {
    const trimmed = (value || "").replace(/\s+/g, " ").trim();
    if (trimmed) items.push(trimmed);
  };

  lines.forEach(line => {
    const matches = Array.from(line.matchAll(inlinePattern));
    if (!matches.length) {
      if (items.length) {
        items[items.length - 1] = `${items[items.length - 1]} ${line}`.trim();
      }
      return;
    }
    if (matches.length === 1 && /^\s*\d+\s*(?:[.)]|-)\s+/.test(line)) {
      pushItem(line.replace(/^\s*\d+\s*(?:[.)]|-)\s+/, ""));
      return;
    }
    matches.forEach((match, idx) => {
      const start = (match.index ?? 0) + match[0].length;
      const end = idx + 1 < matches.length ? (matches[idx + 1].index ?? line.length) : line.length;
      pushItem(line.slice(start, end));
    });
  });

  return items;
}

function extractInlineReferences(answerText) {
  const text = typeof answerText === "string" ? answerText : "";
  if (!text.trim()) return { items: [], startIndex: null };
  const match = /^\s*(references|sources)\s*:\s*(.*)$/im.exec(text);
  if (!match) return { items: [], startIndex: null };

  const startIndex = Number.isFinite(match.index) ? match.index : null;
  const headerRemainder = (match[2] || "").trim();
  let remainder = text.slice((match.index ?? 0) + match[0].length);
  if (remainder.startsWith("\r\n")) remainder = remainder.slice(2);
  else if (remainder.startsWith("\n")) remainder = remainder.slice(1);
  let body = headerRemainder;
  if (remainder) body = body ? `${body}\n${remainder}` : remainder;
  if (!body.trim()) return { items: [], startIndex };

  const noteIndex = body.indexOf(SOURCES_UNAVAILABLE_NOTE);
  if (noteIndex >= 0) body = body.slice(0, noteIndex);
  const blankMatch = body.match(/\n\s*\n/);
  if (blankMatch?.index != null) body = body.slice(0, blankMatch.index);

  const items = parseInlineReferenceItems(body);
  return { items, startIndex };
}

function extractReferenceUrls(block) {
  const entries = [];
  const urlRegex = /https?:\/\/[^\s<>\)]+/gi;
  block.split(/\n/).forEach(line => {
    const matches = Array.from(line.matchAll(urlRegex));
    matches.forEach(match => {
      const url = normalizeCitationUrl(match[0]);
      if (isLikelyValidCitationUrl(url)) {
        entries.push({ url, text: line.trim() || url });
      }
    });
  });
  return entries;
}

function deriveConversationTitle(messages) {
  const firstUser = messages.find(msg => msg.role === "user");
  const source = (firstUser?.text || messages[0]?.text || "Conversation").trim();
  if (!source) return "Conversation";
  return source.length > 60 ? `${source.slice(0, 57)}...` : source;
}

function generateMessageId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function logMessage(role, text, meta = {}) {
  const entryText = text || meta.imageAlt || (meta.imageUrl ? "[image]" : "");
  const entry = {
    id: meta.id || generateMessageId(),
    role,
    text: entryText,
    at: meta.timestamp || formatTime(),
    createdAt: meta.createdAt || Date.now(),
    model: meta.model || (role === "assistant" ? currentModel : undefined),
  };
  if (Array.isArray(meta.attachments) && meta.attachments.length) {
    entry.attachments = cloneMessageAttachments(meta.attachments);
  }
  if (meta.imageUrl) {
    entry.imageUrl = meta.imageUrl;
  }
  if (meta.imageAlt) {
    entry.imageAlt = meta.imageAlt;
  }
  if (meta.docId) entry.docId = meta.docId;
  if (meta.docTitle) entry.docTitle = meta.docTitle;
  if (meta.lectureId) entry.lectureId = meta.lectureId;
  if (meta.extractedKey) entry.extractedKey = meta.extractedKey;
  if (meta.requestId) entry.requestId = meta.requestId;
  if (Array.isArray(meta.references)) entry.references = meta.references.slice(0, 50);
  if (Array.isArray(meta.evidence)) entry.evidence = meta.evidence.slice(0, 50);
  if (typeof meta.renderedMarkdown === "string") entry.renderedMarkdown = meta.renderedMarkdown;
  if (Array.isArray(meta.sources)) {
    const normalized = normalizeStoredSources(meta.sources);
    if (normalized.length) entry.sources = normalized;
  }
  if (Array.isArray(meta.citations)) {
    const normalized = normalizeStoredCitations(meta.citations);
    if (normalized.length) entry.citations = normalized;
  }
  if (Array.isArray(meta.answerSegments)) {
    const sanitized = sanitizeAnswerSegmentsForStorage(meta.answerSegments);
    if (sanitized.length) entry.answerSegments = sanitized;
  }
  if (typeof meta.rawPrompt === "string") entry.rawPrompt = meta.rawPrompt;
  if (typeof meta.cleanedPrompt === "string") entry.cleanedPrompt = meta.cleanedPrompt;
  if (Array.isArray(meta.topics)) entry.topics = meta.topics.slice(0, 20);
  upsertConversationMessage(entry);
  persistConversation();
}

function persistConversation() {
  const now = Date.now();
  const overrideTitle = titleOverrides?.[activeConversationId] || "";
  const title = overrideTitle || (conversation.length ? deriveConversationTitle(conversation) : (activeConversationMeta.title || "Conversation"));
  const selectedDocId = activeLibraryDoc?.docId || activeConversationMeta.selectedDocId || null;
  const selectedDocTitle = activeLibraryDoc
    ? (getLectureDisplayLabel(activeLibraryDoc) || activeLibraryDoc.title || activeLibraryDoc.key || activeLibraryDoc.docId || "")
    : (activeConversationMeta.selectedDocTitle || "");
  activeConversationMeta = {
    ...activeConversationMeta,
    id: activeConversationId,
    title,
    createdAt: activeConversationMeta.createdAt || now,
    updatedAt: now,
    selectedDocId,
    selectedDocTitle,
  };
  const entry = {
    ...activeConversationMeta,
    messages: dedupeMessages(
      conversation.map(msg => ({
        ...msg,
        attachments: Array.isArray(msg.attachments) ? cloneMessageAttachments(msg.attachments) : undefined,
      })),
    ),
  };
  saveConversationDocument(entry);
  storedConversations = (storedConversations || []).filter(conv => conv.id !== activeConversationId);
  storedConversations.unshift({
    id: entry.id,
    title: entry.title,
    createdAt: entry.createdAt,
    updatedAt: entry.updatedAt,
    selectedDocId: entry.selectedDocId,
    selectedDocTitle: entry.selectedDocTitle,
  });
  saveStoredConversations();
  saveActiveConversationId(activeConversationId);
  scheduleConversationSync(entry);
  renderConversationList();
}

function shrinkConversationForStorage(conv) {
  if (!conv || typeof conv !== "object") return null;
  const pruned = pruneMessagesForStorage(Array.isArray(conv.messages) ? conv.messages : []);
  const safeMessages = pruned.messages;
  return {
    id: conv.id || generateConversationId(),
    title: conv.title || "Conversation",
    createdAt: typeof conv.createdAt === "number" ? conv.createdAt : Date.now(),
    updatedAt: typeof conv.updatedAt === "number" ? conv.updatedAt : Date.now(),
    selectedDocId: conv.selectedDocId || null,
    selectedDocTitle: conv.selectedDocTitle || "",
    truncated: pruned.truncated || Boolean(conv.truncated),
    messages: safeMessages,
  };
}

function pruneMessagesForStorage(messages) {
  const sanitized = [];
  let totalChars = 0;
  let truncated = false;
  for (let i = messages.length - 1; i >= 0; i -= 1) {
    if (sanitized.length >= MAX_STORED_MESSAGES) {
      truncated = true;
      break;
    }
    const safeMsg = sanitizeMessageForStorage(messages[i]);
    if (!safeMsg) continue;
    const textLen = typeof safeMsg.text === "string" ? safeMsg.text.length : 0;
    if (sanitized.length && totalChars + textLen > MAX_STORED_CHARS) {
      truncated = true;
      break;
    }
    totalChars += textLen;
    sanitized.push(safeMsg);
  }
  return { messages: dedupeMessages(sanitized.reverse()), truncated };
}

function sanitizeAnswerSegmentsForStorage(segments) {
  if (!Array.isArray(segments) || !segments.length) return [];
  const sanitized = [];
  for (const seg of segments) {
    if (!seg || typeof seg !== "object") continue;
    if (seg.type === "text") {
      const text = typeof seg.text === "string" ? seg.text : "";
      if (!text) continue;
      sanitized.push({ type: "text", text });
    } else if (seg.type === "citation") {
      const id = Number(seg.id ?? seg.i ?? seg.index);
      const url = typeof seg.url === "string" ? seg.url : "";
      if (!Number.isFinite(id) || !url) continue;
      const entry = { type: "citation", id, url };
      if (typeof seg.title === "string" && seg.title) entry.title = seg.title;
      sanitized.push(entry);
    }
    if (sanitized.length >= MAX_STORED_SEGMENTS) break;
  }
  return sanitized;
}

function normalizeStoredSources(sources, citationMap) {
  const normalized = normalizeCitedSources(sources, citationMap);
  return normalized.slice(0, MAX_STORED_SOURCES);
}

function normalizeStoredCitations(citations, citationMap) {
  const normalized = normalizeCitedSources(citations, citationMap);
  return normalized.slice(0, MAX_STORED_CITATIONS);
}

function sanitizeMessageForStorage(raw) {
  if (!raw || typeof raw !== "object") return null;
  const role = raw.role === "assistant" ? "assistant" : raw.role === "system" ? "system" : "user";
  const text = typeof raw.text === "string" ? raw.text : "";
  const finalText = clampTextForStorage(text);
  const msg = {
    id: typeof raw.id === "string" ? raw.id : generateMessageId(),
    role,
    text: finalText,
  };
  if (typeof raw.at === "string") msg.at = raw.at;
  msg.createdAt = typeof raw.createdAt === "number" ? raw.createdAt : Date.now();
  if (typeof raw.model === "string") msg.model = raw.model;
  if (typeof raw.imageAlt === "string") msg.imageAlt = raw.imageAlt;
  if (raw.docId) msg.docId = raw.docId;
  if (raw.docTitle) msg.docTitle = raw.docTitle;
  if (raw.lectureId) msg.lectureId = raw.lectureId;
  if (raw.extractedKey) msg.extractedKey = raw.extractedKey;
  if (raw.requestId) msg.requestId = raw.requestId;
  if (Array.isArray(raw.references)) msg.references = raw.references.slice(0, 50);
  if (Array.isArray(raw.evidence)) msg.evidence = raw.evidence.slice(0, 50);
  if (typeof raw.renderedMarkdown === "string") {
    const rendered = clampTextForStorage(raw.renderedMarkdown);
    if (rendered && rendered !== msg.text) msg.renderedMarkdown = rendered;
  }
  if (Array.isArray(raw.sources)) {
    const normalized = normalizeStoredSources(raw.sources);
    if (normalized.length) msg.sources = normalized;
  }
  if (Array.isArray(raw.citations)) {
    const normalized = normalizeStoredCitations(raw.citations);
    if (normalized.length) msg.citations = normalized;
  }
  if (Array.isArray(raw.answerSegments)) {
    const sanitized = sanitizeAnswerSegmentsForStorage(raw.answerSegments);
    if (sanitized.length) msg.answerSegments = sanitized;
  }
  if (typeof raw.rawPrompt === "string") msg.rawPrompt = raw.rawPrompt;
  if (typeof raw.cleanedPrompt === "string") msg.cleanedPrompt = raw.cleanedPrompt;
  if (Array.isArray(raw.topics)) msg.topics = raw.topics.slice(0, 20);
  const cleanedImage = cleanPreviewUrl(raw.imageUrl);
  if (cleanedImage) msg.imageUrl = cleanedImage;
  if (Array.isArray(raw.attachments)) {
    const safeAttachments = raw.attachments
      .map(sanitizeAttachmentForStorage)
      .filter(Boolean);
    if (safeAttachments.length) {
      msg.attachments = safeAttachments;
    }
  }
  return msg;
}

function clampTextForStorage(text) {
  if (typeof text !== "string") return "";
  if (text.length <= MAX_MESSAGE_CHARS) return text;
  return `${text.slice(0, MAX_MESSAGE_CHARS)}${TRUNCATION_NOTICE}`;
}

function cleanPreviewUrl(value) {
  if (typeof value !== "string" || !value.trim()) return "";
  const trimmed = value.trim();
  if (trimmed.startsWith("data:")) return "";
  return trimmed;
}

function sanitizeAttachmentForStorage(att) {
  const normalized = normalizeMessageAttachment(att);
  if (!normalized) return null;
  const previewUrl = cleanPreviewUrl(normalized.previewUrl);
  return {
    filename: normalized.filename,
    url: normalized.url,
    bucket: normalized.bucket,
    key: normalized.key,
    textKey: normalized.textKey,
    ocrStatus: normalized.ocrStatus,
    ocrWarning: normalized.ocrWarning,
    visionWarning: normalized.visionWarning,
    fileId: normalized.fileId,
    visionFileId: normalized.visionFileId,
    mimeType: normalized.mimeType,
    previewUrl: previewUrl || undefined,
  };
}

function dedupeMessages(list) {
  const result = [];
  let prev = null;
  for (const msg of list || []) {
    const text = typeof msg?.text === "string" ? msg.text : typeof msg?.content === "string" ? msg.content : "";
    const key = msg.id ? `id:${msg.id}` : `${msg.role}|${text}|${msg.imageUrl || ""}|${(msg.attachments || []).length}`;
    if (!prev || prev !== key) {
      result.push(msg);
      prev = key;
    }
  }
  return result;
}

function upsertConversationMessage(msg) {
  if (!msg || !msg.id) return null;
  const idx = conversation.findIndex(m => m.id === msg.id);
  if (idx >= 0) {
    conversation[idx] = { ...conversation[idx], ...msg };
    return conversation[idx];
  }
  conversation.push(msg);
  return msg;
}

function getConversationTitle(conv) {
  if (!conv || !conv.id) return "Conversation";
  const override = titleOverrides?.[conv.id];
  return (override || conv.title || "Conversation").trim() || "Conversation";
}

function setConversationTitleOverride(id, title, opts = {}) {
  if (!id) return;
  if (!title) {
    delete titleOverrides[id];
  } else {
    titleOverrides[id] = title;
  }
  saveTitleOverrides(titleOverrides);
  const skipSync = Boolean(opts.skipSync);
  const skipPersist = Boolean(opts.skipPersist);
  if (!skipSync) {
    const entry = storedConversations.find(conv => conv.id === id);
    if (entry) {
      entry.title = title || entry.title || "Conversation";
      entry.updatedAt = Date.now();
      saveStoredConversations();
      renderConversationList();
    }
    const doc = loadConversationDocument(id);
    if (doc) {
      doc.title = title || doc.title || "Conversation";
      doc.updatedAt = Date.now();
      saveConversationDocument(doc);
      scheduleConversationSync(doc);
    }
  }
  if (!skipPersist && id === activeConversationId) {
    persistConversation();
  }
}

let persistConversationTimer = null;
function schedulePersistConversation(delay = 400) {
  if (persistConversationTimer) return;
  persistConversationTimer = setTimeout(() => {
    persistConversationTimer = null;
    persistConversation();
  }, delay);
}

function renderConversationList() {
  if (!savedList) return;
  const query = (conversationSearch?.value || "").trim().toLowerCase();
  storedConversations.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  const visible = query
    ? storedConversations.filter(conv => getConversationTitle(conv).toLowerCase().includes(query))
    : storedConversations;
  savedList.innerHTML = "";
  if (!visible.length) {
    savedPlaceholder.hidden = false;
    if (savedPlaceholder) {
      savedPlaceholder.textContent = query ? "No conversations match your search." : "No saved conversations yet.";
    }
    return;
  }
  savedPlaceholder.hidden = true;
  visible.forEach(conv => {
    const li = document.createElement("li");
    li.className = "saved-item";
    li.dataset.convId = conv.id;
    li.style.position = "relative";
    if (conv.id === activeConversationId) li.classList.add("active");

    const main = document.createElement("button");
    main.type = "button";
    main.className = "saved-item-main";
    main.addEventListener("click", () => loadConversation(conv.id));

    if (editingConversationId === conv.id) {
      const input = document.createElement("input");
      input.type = "text";
      input.className = "conversation-search";
      input.value = getConversationTitle(conv);
      const save = document.createElement("button");
      save.type = "button";
      save.className = "btn btn-primary";
      save.textContent = "Save";
      const cancel = document.createElement("button");
      cancel.type = "button";
      cancel.className = "btn btn-ghost";
      cancel.textContent = "Cancel";
      const controls = document.createElement("div");
      controls.style.display = "flex";
      controls.style.gap = "6px";
      controls.append(save, cancel);
      main.append(input);
      li.append(main, controls);
      setTimeout(() => input.focus(), 0);
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          save.click();
        } else if (event.key === "Escape") {
          editingConversationId = null;
          renderConversationList();
        }
      });
      save.addEventListener("click", () => {
        const nextTitle = input.value.trim() || "Conversation";
        setConversationTitleOverride(conv.id, nextTitle);
        storedConversations = storedConversations.map(entry => (
          entry.id === conv.id ? { ...entry, title: nextTitle } : entry
        ));
        if (activeConversationId === conv.id) {
          activeConversationMeta = { ...activeConversationMeta, title: nextTitle };
        }
        saveStoredConversations();
        const doc = loadConversationDocument(conv.id);
        if (doc) {
          doc.title = nextTitle;
          saveConversationDocument(doc);
        }
        editingConversationId = null;
        renderConversationList();
      });
      cancel.addEventListener("click", () => {
        editingConversationId = null;
        renderConversationList();
      });
      savedList.append(li);
      return;
    }

    const title = document.createElement("div");
    title.className = "saved-item-title";
    title.textContent = getConversationTitle(conv);
    const meta = document.createElement("div");
    meta.className = "saved-item-meta";
    meta.textContent = formatRelativeTime(conv.updatedAt || Date.now());
    main.append(title, meta);

    const actions = document.createElement("div");
    actions.className = "saved-item-actions";
    const menuBtn = document.createElement("button");
    menuBtn.type = "button";
    menuBtn.className = "menu-trigger";
    menuBtn.setAttribute("aria-label", "Conversation actions");
    menuBtn.setAttribute("aria-haspopup", "menu");
    menuBtn.setAttribute("aria-expanded", String(openConversationMenuId === conv.id));
    menuBtn.title = "Conversation actions";
    menuBtn.textContent = "...";
    menuBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      const next = openConversationMenuId === conv.id ? null : conv.id;
      openConversationMenuId = next;
      if (next !== conv.id) {
        openExportMenuId = null;
        pendingDeleteId = null;
      }
      renderConversationList();
    });

    const menu = document.createElement("div");
    menu.className = "menu-popover";
    menu.setAttribute("role", "menu");
    if (openConversationMenuId === conv.id) menu.classList.add("is-open");

    if (pendingDeleteId === conv.id) {
      const confirm = document.createElement("div");
      confirm.className = "menu-item";
      confirm.textContent = "Delete this conversation?";
      const confirmBtn = document.createElement("button");
      confirmBtn.type = "button";
      confirmBtn.className = "btn btn-primary";
      confirmBtn.textContent = "Delete";
      const cancelBtn = document.createElement("button");
      cancelBtn.type = "button";
      cancelBtn.className = "btn btn-ghost";
      cancelBtn.textContent = "Cancel";
      const row = document.createElement("div");
      row.style.display = "flex";
      row.style.gap = "6px";
      row.append(confirmBtn, cancelBtn);
      menu.append(confirm, row);
      confirmBtn.addEventListener("click", () => {
        deleteConversation(conv.id);
        pendingDeleteId = null;
        openConversationMenuId = null;
      });
      cancelBtn.addEventListener("click", () => {
        pendingDeleteId = null;
        renderConversationList();
      });
    } else {
      const rename = document.createElement("button");
      rename.type = "button";
      rename.className = "menu-item";
      rename.textContent = "Rename";
      rename.addEventListener("click", () => {
        editingConversationId = conv.id;
        openConversationMenuId = null;
        renderConversationList();
      });

      const exportBtn = document.createElement("button");
      exportBtn.type = "button";
      exportBtn.className = "menu-item";
      exportBtn.textContent = "Export";
      exportBtn.addEventListener("click", () => {
        openExportMenuId = openExportMenuId === conv.id ? null : conv.id;
        renderConversationList();
      });

      const del = document.createElement("button");
      del.type = "button";
      del.className = "menu-item";
      del.textContent = "Delete";
      del.addEventListener("click", () => {
        pendingDeleteId = conv.id;
        renderConversationList();
      });

      menu.append(rename, exportBtn);
      if (openExportMenuId === conv.id) {
        const exportRow = document.createElement("div");
        exportRow.style.display = "flex";
        exportRow.style.flexDirection = "column";
        exportRow.style.gap = "6px";
        const jsonBtn = document.createElement("button");
        jsonBtn.type = "button";
        jsonBtn.className = "menu-item";
        jsonBtn.textContent = "Download JSON";
        jsonBtn.addEventListener("click", () => exportConversation(conv.id, "json"));
        const mdBtn = document.createElement("button");
        mdBtn.type = "button";
        mdBtn.className = "menu-item";
        mdBtn.textContent = "Download Markdown";
        mdBtn.addEventListener("click", () => exportConversation(conv.id, "markdown"));
        exportRow.append(jsonBtn, mdBtn);
        menu.append(exportRow);
      }
      menu.append(del);
    }

    actions.append(menuBtn, menu);
    li.append(main, actions);
    savedList.append(li);
  });
}

function applyConversationLoadedState(id, loaded, entry) {
  activeConversationId = id;
  const resolvedTitle = getConversationTitle({ id, title: loaded.title || entry?.title || "Conversation" });
  activeConversationMeta = {
    id,
    title: resolvedTitle,
    createdAt: loaded.createdAt || Date.now(),
    updatedAt: loaded.updatedAt || Date.now(),
    selectedDocId: loaded.selectedDocId || entry?.selectedDocId || null,
    selectedDocTitle: loaded.selectedDocTitle || entry?.selectedDocTitle || "",
  };
  if (!entry) {
    storedConversations = storedConversations.filter(conv => conv.id !== id);
    storedConversations.unshift({
      id,
      title: activeConversationMeta.title,
      createdAt: activeConversationMeta.createdAt,
      updatedAt: activeConversationMeta.updatedAt,
      selectedDocId: activeConversationMeta.selectedDocId,
      selectedDocTitle: activeConversationMeta.selectedDocTitle,
    });
  } else {
    entry.title = activeConversationMeta.title;
    entry.createdAt = activeConversationMeta.createdAt;
    entry.updatedAt = activeConversationMeta.updatedAt;
    entry.selectedDocId = activeConversationMeta.selectedDocId;
    entry.selectedDocTitle = activeConversationMeta.selectedDocTitle;
  }
  saveStoredConversations();
  saveActiveConversationId(id);
  chatLog.innerHTML = "";
  const messages = dedupeMessages(loaded.messages || []);
  if (!messages.length) {
    chatLog.dataset.empty = "true";
  } else {
    messages.forEach(msg => {
      const bubble = appendChatMessage(msg.role, msg.text, {
        track: false,
        model: msg.model,
        attachments: Array.isArray(msg.attachments) ? msg.attachments : undefined,
        imageUrl: msg.imageUrl,
        imageAlt: msg.imageAlt,
        msgId: msg.id,
        createdAt: msg.createdAt,
        docId: msg.docId || msg.lectureId || null,
        requestId: msg.requestId,
        references: msg.references,
        evidence: msg.evidence,
        extractedKey: msg.extractedKey,
        docTitle: msg.docTitle,
        sources: msg.sources,
        citations: msg.citations,
        answerSegments: msg.answerSegments,
        renderedMarkdown: msg.renderedMarkdown,
      });
      if (msg.docId || msg.lectureId) {
        bubble.dataset.docId = msg.docId || msg.lectureId;
      }
    });
    chatLog.dataset.empty = "false";
  }
  if (DEBUG_PERSIST) {
    console.log("[PERSIST_LOAD]", { conv: id, messages: messages.length });
  }
  conversation.length = 0;
  messages.forEach(msg => conversation.push({ ...msg }));
  const lastUser = [...messages].reverse().find(msg => msg.role === "user");
  if (lastUser) {
    if (insightsLog) insightsLog.textContent = `Last ask: ${lastUser.text}`;
    updateHashtagCloud(lastUser.text);
  } else {
    if (insightsLog) insightsLog.textContent = "Conversation loaded.";
    if (hashtagCloud) hashtagCloud.innerHTML = "";
  }
  if (activeConversationMeta.selectedDocId) {
    setLibrarySelection({
      docId: activeConversationMeta.selectedDocId,
      title: activeConversationMeta.selectedDocTitle || activeConversationMeta.selectedDocId,
      status: "ready",
    });
  } else {
    setLibrarySelection(null);
  }
  renderMathIfReady(chatLog);
  renderConversationList();
  scrollToHash({ behavior: "auto" });
}

async function loadConversation(id, opts = {}) {
  if (id && id !== activeConversationId) {
    persistConversation();
  }
  const entry = storedConversations.find(conv => conv.id === id);
  const localDoc = normalizeConversationDocument(loadConversationDocument(id));
  const loadToken = ++conversationLoadToken;
  if (!entry && !localDoc) {
    const remoteDoc = await fetchConversationDocumentFromServer(id);
    if (loadToken !== conversationLoadToken) return;
    if (remoteDoc) {
      saveConversationDocument(remoteDoc);
      applyConversationLoadedState(id, remoteDoc, entry);
      return;
    }
    showConversationWarning(id, "This conversation could not be loaded yet. Try again.");
    return;
  }
  const loaded = localDoc || {
    id,
    messages: [],
    title: entry?.title || "Conversation",
    createdAt: entry?.createdAt || Date.now(),
    updatedAt: entry?.updatedAt || Date.now(),
    selectedDocId: entry?.selectedDocId || null,
    selectedDocTitle: entry?.selectedDocTitle || "",
  };
  applyConversationLoadedState(id, loaded, entry);

  const entryUpdatedAt = entry?.updatedAt || 0;
  const localUpdatedAt = localDoc?.updatedAt || 0;
  const shouldFetchRemote = Boolean(opts.forceRemote)
    || !localDoc
    || !(localDoc.messages || []).length
    || entryUpdatedAt > localUpdatedAt;
  if (!shouldFetchRemote) return;
  const remoteDoc = await fetchConversationDocumentFromServer(id);
  if (loadToken !== conversationLoadToken) return;
  if (remoteDoc) {
    if (!remoteDoc.messages.length && localDoc?.messages?.length) {
      scheduleConversationSync(localDoc);
      return;
    }
    const merged = mergeConversationDocuments(localDoc, remoteDoc);
    let hasChanges = false;
    if (merged) {
      if (merged.updatedAt !== activeConversationMeta.updatedAt) hasChanges = true;
      if (merged.messages.length !== conversation.length) {
        hasChanges = true;
      } else if (merged.messages.length) {
        const lastMerged = merged.messages[merged.messages.length - 1];
        const lastLocal = conversation[conversation.length - 1];
        if (lastMerged?.id !== lastLocal?.id || lastMerged?.text !== lastLocal?.text) {
          hasChanges = true;
        }
      }
    }
    if (merged) {
      saveConversationDocument(merged);
    }
    if (merged && hasChanges) {
      applyConversationLoadedState(id, merged, entry);
    }
    if (!merged?.messages?.length) {
      showConversationWarning(id, "This conversation has no stored transcript yet.");
    }
  } else if (localDoc?.messages?.length) {
    scheduleConversationSync(localDoc);
  } else {
    showConversationWarning(id, "This conversation could not be loaded yet. Try again.");
  }
}

function startNewConversation() {
  if (conversation.length) persistConversation();
  if (persistConversationTimer) {
    clearTimeout(persistConversationTimer);
    persistConversationTimer = null;
  }
  conversationLoadToken += 1;
  activeConversationId = generateConversationId();
  setConversationTitleOverride(activeConversationId, "", { skipSync: true, skipPersist: true });
  activeConversationMeta = {
    id: activeConversationId,
    title: "Conversation",
    createdAt: Date.now(),
    updatedAt: Date.now(),
    selectedDocId: null,
    selectedDocTitle: "",
  };
  saveActiveConversationId(activeConversationId);
  conversation.length = 0;
  activePdfSession = null;
  clearOcrSession();
  setLibrarySelection(null);
  if (chatLog) {
    chatLog.innerHTML = "";
    chatLog.dataset.empty = "true";
  }
  hideCitationTooltip();
  if (insightsLog) insightsLog.textContent = "No topics logged yet. Start chatting to see trends.";
  if (hashtagCloud) hashtagCloud.innerHTML = "";
  renderConversationList();
}

function getHistoryPayload(limit = 10) {
  return conversation
    .filter(entry => entry && (entry.role === "user" || entry.role === "assistant") && typeof entry.text === "string")
    .slice(-limit)
    .map(entry => ({ role: entry.role, text: entry.text }))
    .filter(entry => entry.text && entry.text.trim().length > 0);
}

async function summarizeAttachment(fileId, filename) {
  if (!fileId || attachmentSummaries[fileId]) return attachmentSummaries[fileId];
  const task = (async () => {
    try {
      const response = await fetch("/api/extract", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ file_id: fileId, filename }),
      });
      if (!response.ok) {
        const errText = await response.text().catch(() => "");
        throw new Error(errText || "Could not summarize document.");
      }
      const payload = await response.json().catch(() => ({}));
      attachmentSummaries[fileId] = payload;
      return payload;
    } catch (error) {
      console.warn("Attachment summary failed", error);
      attachmentSummaries[fileId] = null;
      return null;
    }
  })();
  documentSummaryTasks.add(task);
  try {
    return await task;
  } finally {
    documentSummaryTasks.delete(task);
  }
}

function extractKeywords(text) {
  return new Set(
    (text || "")
      .toLowerCase()
      .match(/\b[a-z0-9]{4,}\b/g)
      ?.filter(word => !STOPWORDS.has(word)) || [],
  );
}

function shouldAttachHistory(prompt, historyEntries) {
  if (!historyEntries || !historyEntries.length) return false;
  if (FOLLOWUP_REGEX.test(prompt)) return true;
  const promptKeywords = extractKeywords(prompt);
  if (!promptKeywords.size) return false;
  const historyKeywords = new Set();
  historyEntries.forEach(entry => {
    extractKeywords(entry.text).forEach(word => historyKeywords.add(word));
  });
  let overlap = 0;
  promptKeywords.forEach(word => {
    if (historyKeywords.has(word)) overlap += 1;
  });
  if (overlap >= 2) return true;
  const lastUser = [...historyEntries].reverse().find(entry => entry.role === "user");
  if (!lastUser) return false;
  const lastKeywords = extractKeywords(lastUser.text);
  let lastOverlap = 0;
  promptKeywords.forEach(word => {
    if (lastKeywords.has(word)) lastOverlap += 1;
  });
  return lastOverlap >= 1 && promptKeywords.size <= 4;
}

const MATH_TOKENS = /\\(frac|sqrt|sum|int|pi|eta|theta|delta|Delta|alpha|beta|gamma|lambda|mu|nu|omega|phi|psi|times|leq|geq|approx|neq|pm)\b|[∑√πΔΩαβγδηθλμνξπρστυφχψω]/i;
function needsMathWrapping(raw) {
  if (!raw) return false;
  if (/\\\[|\\\(|\$\$/.test(raw)) return false;
  return MATH_TOKENS.test(raw);
}

function wrapMath(html) {
  return `\\[${html}\\]`;
}

function formatLine(line) {
  let raw = line || "";
  raw = raw.replace(/^([A-Za-z][^:\n]{1,40}):\s+(.+)/, (_, label, rest) => `**${label.trim()}**: ${rest.trim()}`);
  let html = escapeHTML(raw);
  html = html.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/`([^`]+)`/g, "<code>$1</code>");
  html = html.replace(/\[(C\d+)\]/gi, (_, id) => `<span class="cite-badge">${escapeHTML(id)}</span>`);
  html = html.replace(/\[(\d+)\]/g, (_, id) => `<sup class="citation-sup" data-cite-inline="true"><a class="cite citation-chip" data-cite-id="${id}" data-cite-inline="true" aria-disabled="true">[${id}]</a></sup>`);
  if (needsMathWrapping(raw)) {
    html = wrapMath(html);
  }
  return html;
}

function renderTableCaptionNode(text, { usedIds, idPrefix } = {}) {
  const trimmed = (text || "").trim();
  if (!trimmed) return null;
  const headingMatch = /^(#{2,6})\s+(.*)$/.exec(trimmed);
  if (headingMatch) {
    const level = headingMatch[1].length;
    const title = headingMatch[2].trim();
    let tag = "h4";
    if (level <= 2) tag = "h2";
    else if (level === 3) tag = "h3";
    const heading = document.createElement(tag);
    heading.className = "section-heading";
    heading.innerHTML = formatLine(title);
    if (usedIds) {
      heading.id = buildSectionId(title, usedIds.size, usedIds, { prefix: idPrefix });
    }
    return heading;
  }
  const caption = document.createElement("p");
  caption.className = "table-caption";
  caption.innerHTML = formatLine(trimmed);
  return caption;
}

function pickTableTone(headers, rows) {
  const blob = `${headers.join(" ")} ${rows.flat().join(" ")}`.toLowerCase();
  if (/\b(clinic|patient|symptom|diagnos|treat|therapy|drug|dose|respiratory|patholog)/.test(blob)) return "clinical";
  if (/\b(timeline|phase|milestone|date|deadline|sprint|week|month)/.test(blob)) return "timeline";
  if (/\b(risk|issue|mitigation|hazard|severity)/.test(blob)) return "risk";
  if (/\b(metric|kpi|score|roi|budget|cost|revenue|impact)/.test(blob)) return "metric";
  if (/\b(compare|versus|option|variant|pros|cons|choice|alt)/.test(blob)) return "comparison";
  return "default";
}

function buildMarkdownTable(block) {
  return renderMarkdownTable(block, {
    formatCell: formatLine,
    pickTone: pickTableTone,
  });
}

function buildHtmlTable(block) {
  return renderHtmlTable(block, {
    formatCell: formatLine,
    pickTone: pickTableTone,
  });
}

function splitHtmlTableBlocks(text) {
  if (!text || !/<table[\s\S]*?>/i.test(text)) return null;
  const parts = [];
  const regex = /<table[\s\S]*?<\/table>/gi;
  let lastIndex = 0;
  let match = null;
  while ((match = regex.exec(text))) {
    if (match.index > lastIndex) {
      parts.push({ type: "text", text: text.slice(lastIndex, match.index) });
    }
    parts.push({ type: "table", html: match[0] });
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < text.length) {
    parts.push({ type: "text", text: text.slice(lastIndex) });
  }
  return parts.length ? parts : null;
}

function extractTrailingCaption(text) {
  if (!text) return { text: "", caption: "" };
  const lines = text.split(/\n/);
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    const line = (lines[i] || "").trim();
    if (!line) continue;
    const match = /^caption\s*:\s*(.+)$/i.exec(line);
    if (!match) return { text, caption: "" };
    lines.splice(i, 1);
    return { text: lines.join("\n"), caption: match[1].trim() };
  }
  return { text, caption: "" };
}

function renderTableFallback(block) {
  const pre = document.createElement("pre");
  pre.className = "table-fallback";
  pre.style.whiteSpace = "pre-wrap";
  pre.textContent = block;
  return pre;
}

function findHeadingBefore(node) {
  let prev = node.previousElementSibling;
  while (prev) {
    if (/^H[1-4]$/i.test(prev.tagName)) {
      const txt = (prev.textContent || "").trim();
      if (txt) return txt;
    }
    if (prev.textContent && prev.textContent.trim()) break;
    prev = prev.previousElementSibling;
  }
  return "";
}

function enhanceAutoTables(root) {
  if (!root || !root.querySelectorAll) return;
  const lists = Array.from(root.querySelectorAll("ul,ol"));
  const THRESHOLD_RE = /^(\~?\s*\d+(?:\.\d+)?\s*[–-]\s*\d+(?:\.\d+)?|[<>≥≤]\s*\d+(?:\.\d+)?(?:\s*[a-zA-Z/%]+)?)\s*[:\-–]\s+(.*)$/i;

  lists.forEach(list => {
    if (!list.children || list.children.length < 3) return;
    const items = Array.from(list.children).filter(li => li.tagName === "LI");
    if (items.length < 3) return;
    const parsed = [];
    const refSet = new Set();

    for (const li of items) {
      if (li.querySelector("ul,ol")) return; // nested list: skip
      const text = (li.textContent || "").trim();
      const match = THRESHOLD_RE.exec(text);
      if (!match) return;
      let threshold = match[1].replace(/\s+/g, " ").trim();
      let desc = match[2].trim();
      desc = desc.replace(/\s{2,}/g, " ");
      desc = desc.replace(/\((Slide\s*\d+[^)]*|Page\s*\d+[^)]*)\)/gi, (m) => {
        refSet.add(m.replace(/[()]/g, "").trim());
        return "";
      }).trim();
      parsed.push({ threshold, desc });
    }

    if (parsed.length < 3) return;

    const units = new Set();
    parsed.forEach(({ threshold }) => {
      const unitMatch = threshold.match(/\b([a-zA-Z/%]+)\s*$/);
      if (unitMatch) units.add(unitMatch[1].toLowerCase());
    });
    if (units.size > 1) return;

    const title = findHeadingBefore(list) || "Clinical thresholds/steps";

    const table = document.createElement("table");
    table.className = "owen-table AnswerTable AnswerTable--kv";
    table.dataset.answerTable = "true";
    table.style.setProperty("--answer-table-columns", "2");
    table.setAttribute("data-auto-table-title", title);
    if (refSet.size) {
      table.setAttribute("data-ref-chips", JSON.stringify(Array.from(refSet.values()).slice(0, 12)));
    }
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    ["Range / Threshold", "Findings / Changes"].forEach(label => {
      const th = document.createElement("th");
      th.textContent = label;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    const tbody = document.createElement("tbody");
    parsed.forEach(({ threshold, desc }) => {
      const tr = document.createElement("tr");
      const tdRange = document.createElement("td");
      tdRange.textContent = threshold;
      const tdDesc = document.createElement("td");
      tdDesc.textContent = desc || " ";
      tr.append(tdRange, tdDesc);
      tbody.appendChild(tr);
    });
    table.append(thead, tbody);

    list.replaceWith(table);
  });
}

function enhanceTablesIn(root) {
  if (!root || !root.querySelectorAll) return;
  const tables = Array.from(root.querySelectorAll("table"));
  tables.forEach(table => {
    if (table.closest(".owen-table-card")) return;
    const parent = table.parentElement;
    if (!parent || !parent.contains(table)) return;

    const headerCells = Array.from(table.querySelectorAll("thead th")).map(th => (th.textContent || "").trim());
    const firstRowCells = table.querySelector("tbody tr")?.children.length || 0;
    const columnCount = headerCells.length || firstRowCells;
    table.classList.add("AnswerTable");
    table.dataset.answerTable = "true";
    if (columnCount) table.style.setProperty("--answer-table-columns", String(columnCount));
    if (columnCount === 2) table.classList.add("AnswerTable--kv");

    const titleGuess =
      table.getAttribute("data-title") ||
      table.getAttribute("data-auto-table-title") ||
      (table.querySelector("thead th") ? "Table" : "Auto-generated summary");

    const card = document.createElement("div");
    card.className = "owen-table-card";
    card.dataset.owenTable = "true";

    const header = document.createElement("div");
    header.className = "owen-table-card__header";
    const pill = document.createElement("span");
    pill.className = "owen-pill owen-pill--table";
    pill.textContent = "📊 Table";
    const title = document.createElement("span");
    title.className = "owen-table-card__title";
    title.textContent = titleGuess || "Table";
    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "owen-table-toggle";
    toggle.dataset.wrapToggle = "true";
    header.append(pill, title, toggle);

    const resizer = document.createElement("div");
    resizer.className = "owen-table-resizer";
    const scroller = document.createElement("div");
    scroller.className = "owen-table-scroll";
    const handle = document.createElement("div");
    handle.className = "owen-table-handle";
    handle.setAttribute("aria-hidden", "true");

    const pretty = document.createElement("div");
    pretty.className = "owen-pretty-table";
    const tabs = document.createElement("div");
    tabs.className = "owen-pretty-table__tabs";
    const palette = ["#60a5fa", "#34d399", "#fbbf24", "#f472b6", "#a78bfa", "#38bdf8"];
    headerCells.forEach((label, idx) => {
      const tab = document.createElement("span");
      tab.className = "owen-pretty-table__tab";
      tab.textContent = label || `Column ${idx + 1}`;
      tab.style.background = `linear-gradient(135deg, ${palette[idx % palette.length]}33, ${palette[(idx + 1) % palette.length]}55)`;
      tabs.appendChild(tab);
    });
    const wrap = document.createElement("div");
    wrap.className = "owen-pretty-table__wrap AnswerTableWrap";
    wrap.appendChild(table);
    if (headerCells.length) {
      pretty.appendChild(tabs);
    }
    pretty.appendChild(wrap);

    scroller.appendChild(pretty);
    resizer.append(scroller, handle);
    card.append(header, resizer);

    const refChips = table.getAttribute("data-ref-chips");
    if (refChips) {
      try {
        const parsed = JSON.parse(refChips);
        if (Array.isArray(parsed) && parsed.length) {
          const chipWrap = document.createElement("div");
          chipWrap.className = "reference-chips";
          parsed.slice(0, 12).forEach(text => {
            if (!text || typeof text !== "string") return;
            const chip = document.createElement("span");
            chip.className = "reference-chip";
            chip.textContent = text.slice(0, 180);
            chipWrap.appendChild(chip);
          });
          card.appendChild(chipWrap);
        }
      } catch {
        // ignore malformed ref chips
      }
    }

    safeReplaceNode(table, card, "table-enhance");

    const applyWrapState = (wrapOn) => {
      if (wrapOn) {
        card.classList.remove("owen-table-card--nowrap");
        toggle.dataset.state = "on";
        toggle.textContent = "Wrap: On";
        toggle.setAttribute("aria-pressed", "true");
      } else {
        card.classList.add("owen-table-card--nowrap");
        toggle.dataset.state = "off";
        toggle.textContent = "Wrap: Off";
        toggle.setAttribute("aria-pressed", "false");
      }
      try {
        localStorage.setItem(TABLE_WRAP_PREF_KEY, wrapOn ? "on" : "off");
      } catch {}
    };

    const preferred = (() => {
      try {
        return localStorage.getItem(TABLE_WRAP_PREF_KEY);
      } catch {
        return null;
      }
    })();
    applyWrapState(preferred !== "off");

    toggle.addEventListener("click", () => {
      const isOn = toggle.dataset.state !== "off";
      applyWrapState(!isOn);
    });

    table.addEventListener("mouseover", (e) => {
      const cell = e.target.closest("td,th");
      if (!cell) return;
      const colIndex = Array.from(cell.parentElement?.children || []).indexOf(cell);
      if (colIndex < 0) return;
      const rows = table.querySelectorAll("tr");
      rows.forEach(row => {
        const cells = Array.from(row.children);
        if (cells[colIndex]) cells[colIndex].classList.add("owen-col-hover");
      });
    });
    table.addEventListener("mouseout", (e) => {
      if (!e.target.closest("td,th")) return;
      table.querySelectorAll(".owen-col-hover").forEach(cell => cell.classList.remove("owen-col-hover"));
    });
  });
}

function renderMarkdownContent(blockText) {
  const fragment = document.createDocumentFragment();
  const lines = (blockText || "").split("\n");
  let currentList = null;
  let currentCode = null;
  let paragraph = [];

  const flushParagraph = () => {
    if (!paragraph.length) return;
    const p = document.createElement("p");
    p.innerHTML = formatLine(paragraph.join(" "));
    fragment.appendChild(p);
    paragraph = [];
  };

  const flushList = () => {
    if (!currentList) return;
    fragment.appendChild(currentList.node);
    currentList = null;
  };

  const flushCode = () => {
    if (!currentCode) return;
    const pre = document.createElement("pre");
    const codeEl = document.createElement("code");
    codeEl.textContent = currentCode.lines.join("\n");
    pre.appendChild(codeEl);
    fragment.appendChild(pre);
    currentCode = null;
  };

  lines.forEach(line => {
    const trimmed = line.trimEnd();
    const fence = /^```/.test(trimmed);
    if (fence) {
      if (currentCode) {
        flushCode();
      } else {
        flushParagraph();
        flushList();
        currentCode = { lines: [] };
      }
      return;
    }
    if (currentCode) {
      currentCode.lines.push(line);
      return;
    }

    if (!trimmed) {
      flushParagraph();
      flushList();
      return;
    }

    const headingMatch = /^(#{1,4})\s+(.*)/.exec(trimmed);
    if (headingMatch) {
      flushParagraph();
      flushList();
      const level = headingMatch[1].length;
      let tag = "h4";
      if (level <= 2) tag = "h2";
      else if (level === 3) tag = "h3";
      const h = document.createElement(tag);
      h.innerHTML = formatLine(headingMatch[2]);
      fragment.appendChild(h);
      return;
    }

    const bulletMatch = /^[-*+]\s+(.+)/.exec(trimmed);
    if (bulletMatch) {
      flushParagraph();
      const text = bulletMatch[1].replace(/^([A-Za-z][^:]{1,40}):\s+(.+)/, "**$1:** $2");
      if (!currentList || currentList.type !== "ul") {
        flushList();
        currentList = { type: "ul", node: document.createElement("ul") };
        currentList.node.className = "section-list";
      }
      const li = document.createElement("li");
      li.innerHTML = formatLine(text);
      currentList.node.appendChild(li);
      return;
    }

    const orderedMatch = /^\d+[.)]\s+(.+)/.exec(trimmed);
    if (orderedMatch) {
      flushParagraph();
      const text = orderedMatch[1].replace(/^([A-Za-z][^:]{1,40}):\s+(.+)/, "**$1:** $2");
      if (!currentList || currentList.type !== "ol") {
        flushList();
        currentList = { type: "ol", node: document.createElement("ol") };
        currentList.node.className = "section-list ordered";
      }
      const li = document.createElement("li");
      li.innerHTML = formatLine(text);
      currentList.node.appendChild(li);
      return;
    }

    paragraph.push(trimmed);
  });

  flushParagraph();
  flushList();
  flushCode();

  if (!fragment.childNodes.length && blockText) {
    const p = document.createElement("p");
    p.innerHTML = formatLine(blockText);
    fragment.appendChild(p);
  }
  return fragment;
}

function pickEmoji(title) {
  const tests = [
    { re: /summary|overview|intro/i, emoji: "🧭" },
    { re: /plan|strategy|steps/i, emoji: "🗺️" },
    { re: /tip|advice|hint/i, emoji: "💡" },
    { re: /warning|risk/i, emoji: "⚠️" },
    { re: /result|insight/i, emoji: "📈" },
    { re: /reference|source/i, emoji: "📚" },
    { re: /biology|medical|therapy|clinical/i, emoji: "🧬" },
    { re: /chem|compound|molecule|pharma/i, emoji: "⚗️" },
    { re: /data|stat|analysis|quant/i, emoji: "📊" },
    { re: /neuro|brain/i, emoji: "🧠" },
    { re: /cardio|heart/i, emoji: "❤️" },
    { re: /lab|protocol|experiment/i, emoji: "🧪" },
    { re: /timeline|schedule/i, emoji: "📆" },
  ];
  const found = tests.find(entry => entry.re.test(title));
  return found ? found.emoji : "✨";
}

function normalizeHeadingCandidate(text) {
  if (!text) return "";
  return text
    .replace(/[:\-]+$/, " ")
    .replace(/[^a-z0-9\s]/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeAnchorPrefix(value) {
  if (!value) return "";
  return String(value)
    .replace(/[^a-z0-9_-]+/gi, "-")
    .replace(/-+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function slugifySectionLabel(label) {
  const cleaned = normalizeHeadingCandidate(label || "");
  if (!cleaned) return "";
  return cleaned
    .toLowerCase()
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function isUsableHeading(text) {
  const normalized = normalizeHeadingCandidate(text);
  if (!normalized) return false;
  const words = normalized.split(" ").filter(Boolean);
  if (words.length < 3) return false;
  const weakStarts = ["i", "you", "we", "do", "does", "did", "can", "could", "would", "should"];
  if (weakStarts.includes(words[0].toLowerCase())) return false;
  return true;
}

function prettifyHeading(text) {
  const normalized = normalizeHeadingCandidate(text);
  if (!normalized) return "";
  return normalized
    .split(" ")
    .slice(0, 8)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(" ");
}

function deriveHeadingLabel(block, fallback, index) {
  const source = `${fallback || ""} ${block}`.toLowerCase().replace(/[:]+$/, "").trim();
  const matched = SECTION_LABEL_WHITELIST_LOWER.find(label => source.includes(label));
  return matched ? matched.replace(/\b\w/g, c => c.toUpperCase()) : "Response";
}

function sanitizeAssistantTextForSections(text) {
  if (!text) return "";
  // Preserve the assistant text as-is while still normalizing table markers so nothing gets dropped.
  return normalizeTablesInText(text);
}

function allowExcerptsFromPrompt(prompt) {
  const lower = (prompt || "").toLowerCase();
  if (!lower) return false;
  const triggers = [
    "quote",
    "show slide",
    "show the slide",
    "show excerpt",
    "show the excerpt",
    "show evidence",
    "excerpts",
    "verbatim",
    "evidence",
    "proof",
    "cite page",
  ];
  return triggers.some(trigger => lower.includes(trigger));
}

function normalizeLibraryAskPayload(payload, allowExcerpts) {
  const base = {
    ok: false,
    answer: "",
    references: [],
    evidence: [],
    error: null,
    debug: null,
    versioned: false,
  };
  if (!payload || typeof payload !== "object") return base;
  if (payload.v === 1) {
    const ok = payload.ok === true;
    const answer = typeof payload.answer === "string" ? payload.answer : "";
    const references = Array.isArray(payload.references) ? payload.references : [];
    const evidence = allowExcerpts && Array.isArray(payload.evidence) ? payload.evidence : [];
    const error = payload.error && typeof payload.error.message === "string" ? payload.error : null;
    const debug = payload.debug || null;
    return { ok, answer, references, evidence, error, debug, versioned: true };
  }
  const answer =
    typeof payload.answer === "string"
      ? payload.answer
      : typeof payload.output_text === "string"
        ? payload.output_text
        : typeof payload.response === "string"
          ? payload.response
          : typeof payload.text === "string"
            ? payload.text
            : "";
  const references = Array.isArray(payload.references) ? payload.references : [];
  const evidence = allowExcerpts && Array.isArray(payload.evidence) ? payload.evidence : [];
  return { ok: Boolean((answer || "").trim()), answer, references, evidence, error: null, debug: null, versioned: false };
}

function isMachineTxtLoadFailure(payload) {
  if (!payload || typeof payload !== "object") return false;
  const error = payload.error;
  if (typeof error === "string") {
    return error === "machine_txt_not_found" || error === "machine_txt_unavailable";
  }
  if (error && typeof error === "object") {
    const code = typeof error.code === "string" ? error.code : "";
    return code === "machine_txt_not_found" || code === "machine_txt_unavailable";
  }
  return false;
}

function isTableIntent(prompt) {
  const lower = (prompt || "").toLowerCase();
  if (!lower) return false;
  return /\b(table|tabulate|columns|make a table|create a table|put in a table|table of)\b/.test(lower);
}

function looksLikeMarkdownTable(s) {
  if (!s) return false;
  const t = s.trim();
  if (t.startsWith("|")) return true;
  return /\n\|.+\|\n\|[-:\s|]+\|/m.test(s);
}

function looksLikeHtmlTable(s) {
  if (!s) return false;
  return /<table[\s\S]*?>/i.test(s);
}

function sanitizeAnswerText(answer, { allowExcerpts = false } = {}) {
  if (!answer) return "";
  // Keep the answer text intact (including tables/headings) while normalizing escaped table markers.
  return normalizeTablesInText(answer).trim();
}

function normalizeTablesInText(text) {
  if (!text) return "";
  const segments = (text || "").split(/(```[\s\S]*?```)/g);
  return segments
    .map(seg => {
      if (/^```/.test(seg)) return seg; // leave code blocks untouched
      const normalized = seg
        .split(/\n/)
        .map(line => {
          if (!line.includes("\\|")) return line;
          const hasUnescapedPipe = /(^|[^\\])\|/.test(line);
          if (!hasUnescapedPipe) {
            return line.replace(/\\\|/g, "|");
          }
          return line;
        })
        .join("\n");
      return expandCollapsedTables(normalized);
    })
    .join("");
}

function expandCollapsedTables(block) {
  if (!block) return "";
  if (block.includes("\n")) return block;
  const normalized = block.replace(/\|\s*\|\s*/g, "|\n|");
  const pipeRowPattern = /\|\s*[^|]+\|\s*[^|]+(?:\|[^|]+)*\|/g;
  const matches = normalized.match(pipeRowPattern);
  if (matches && matches.length >= 2) {
    return matches.join("\n");
  }
  return normalized;
}

function parseSectionsFromLabels(text) {
  const labelRe = new RegExp(`^(${SECTION_LABEL_WHITELIST_LOWER.join("|")}):\\s*(.*)$`, "i");
  const lines = text.split(/\r?\n/);
  const sections = [];
  let current = { label: "Response", level: 2, lines: [] };
  let inCode = false;
  lines.forEach(rawLine => {
    const line = rawLine || "";
    const trimmed = line.trim();
    if (/^```/.test(trimmed)) {
      inCode = !inCode;
      current.lines.push(line);
      return;
    }
    if (!inCode) {
      const match = labelRe.exec(trimmed);
      if (match) {
        if (current.lines.length) {
          sections.push({ ...current, text: current.lines.join("\n").trim() });
        }
        const nextLabel = match[1].replace(/\b\w/g, c => c.toUpperCase());
        const initial = match[2] ? [match[2]] : [];
        current = { label: nextLabel, level: 2, lines: initial };
        return;
      }
    }
    current.lines.push(line);
  });
  if (current.lines.length || !sections.length) {
    sections.push({ ...current, text: current.lines.join("\n").trim() });
  }
  return sections.filter(sec => sec && typeof sec.text === "string");
}

function parseSectionsFromText(text) {
  const headingRe = /^(#{2,3})\s+(.+)$/;
  const markerRe = /^<!--\s*SECTION:\s*(.+?)\s*-->$/i;
  const lines = text.split(/\r?\n/);
  const sections = [];
  let current = { label: "Response", level: 2, lines: [] };
  let inCode = false;
  let sawHeading = false;
  lines.forEach(rawLine => {
    const line = rawLine || "";
    const trimmed = line.trim();
    if (/^```/.test(trimmed)) {
      inCode = !inCode;
      current.lines.push(line);
      return;
    }
    if (!inCode) {
      const markerMatch = markerRe.exec(trimmed);
      if (markerMatch && markerMatch[1].trim()) {
        sawHeading = true;
        if (current.lines.length) {
          sections.push({ ...current, text: current.lines.join("\n").trim() });
        }
        current = { label: markerMatch[1].trim(), level: 2, lines: [] };
        return;
      }
      const headingMatch = headingRe.exec(trimmed);
      if (headingMatch && headingMatch[2].trim()) {
        sawHeading = true;
        if (current.lines.length) {
          sections.push({ ...current, text: current.lines.join("\n").trim() });
        }
        current = { label: headingMatch[2].trim(), level: headingMatch[1].length, lines: [] };
        return;
      }
    }
    current.lines.push(line);
  });
  if (current.lines.length || !sections.length) {
    sections.push({ ...current, text: current.lines.join("\n").trim() });
  }
  if (!sawHeading) {
    return parseSectionsFromLabels(text);
  }
  return sections.filter(sec => sec && typeof sec.text === "string");
}

function isCitationsNoise(block) {
  const trimmed = (block || "").trim().toLowerCase();
  if (!trimmed) return false;
  if (/^:::\s*citations/.test(trimmed)) return true;
  if (/^citations?\b/.test(trimmed)) return true;
  if (/^\[\s*\{/.test(trimmed) && trimmed.includes('"url"')) return true;
  if (/^\{\s*"i"\s*:/.test(trimmed)) return true;
  return false;
}

const RESPONSE_SUMMARY_BULLET_LIMIT = 8;
const SECTION_PREVIEW_CHAR_LIMIT = 240;

function normalizeViewMode(value) {
  return value === "summary" ? "summary" : "full";
}

function normalizeCitationMode() {
  return "inline";
}

function getInitialResponseView() {
  return normalizeViewMode(safeGetItem(RESPONSE_VIEW_KEY));
}

function buildSectionId(label, index, usedIds, opts = {}) {
  const registry = usedIds instanceof Set ? usedIds : new Set();
  const slug = slugifySectionLabel(label);
  const prefix = normalizeAnchorPrefix(opts.prefix || "");
  const base = slug || `section-${index + 1}`;
  let id = prefix ? `${prefix}-${base}` : base;
  let suffix = 2;
  while (registry.has(id)) {
    id = `${id}-${suffix++}`;
  }
  registry.add(id);
  return id;
}

function ensureId(node, prefix = "response-body") {
  if (!node) return null;
  if (node.id) return node.id;
  const id = `${prefix}-${Math.random().toString(36).slice(2, 8)}`;
  node.id = id;
  return id;
}

function sectionHasMeaningfulContent(body) {
  if (!body) return false;
  const blocks = Array.from(body.querySelectorAll("p, ul, ol, table, pre, .table-fallback, .owen-table-card"));
  if (!blocks.length) return false;
  return blocks.some(node => (node.textContent || "").trim().length > 0);
}

function buildSectionPreview(body) {
  const preview = document.createElement("div");
  preview.className = "section-preview";
  if (!body) return preview;
  const listItems = Array.from(body.querySelectorAll("li")).filter(li => (li.textContent || "").trim());
  if (listItems.length) {
    const ul = document.createElement("ul");
    ul.className = "section-preview-list";
    listItems.slice(0, 2).forEach(li => {
      const item = document.createElement("li");
      item.textContent = (li.textContent || "").trim();
      ul.appendChild(item);
    });
    preview.appendChild(ul);
    return preview;
  }
  const text = (body.textContent || "").replace(/\s+/g, " ").trim();
  if (!text) return preview;
  const snippet = text.slice(0, SECTION_PREVIEW_CHAR_LIMIT);
  const p = document.createElement("p");
  p.textContent = text.length > SECTION_PREVIEW_CHAR_LIMIT ? `${snippet}…` : snippet;
  preview.appendChild(p);
  return preview;
}

function addInlineUrlsToCitationMap(citationMap, inlineUrls) {
  if (!(citationMap instanceof Map) || !Array.isArray(inlineUrls) || !inlineUrls.length) return;
  const entries = [];
  citationMap.forEach((value, id) => {
    const url = typeof value === "string" ? value : value?.url;
    if (url && isLikelyValidCitationUrl(url)) entries.push({ id: String(id), url });
  });
  let nextId = entries
    .map(entry => parseInt(entry.id, 10))
    .filter(n => !Number.isNaN(n))
    .reduce((max, n) => Math.max(max, n), entries.length ? 1 : 0) + 1;
  const seenUrls = new Set(entries.map(entry => entry.url));
  inlineUrls.forEach((url) => {
    const normalized = normalizeCitationUrl(url);
    if (!normalized || !isLikelyValidCitationUrl(normalized) || seenUrls.has(normalized)) return;
    const id = String(nextId++);
    citationMap.set(id, { url: normalized, title: "", snippet: "", domain: "" });
    seenUrls.add(normalized);
  });
}

function getContextLine() {
  const lecture = getLectureDisplayLabel(activeLibraryDoc) || activeConversationMeta?.selectedDocTitle || "--";
  const materialsTotal = attachments.length + (oneShotFile ? 1 : 0);
  const materialsCount = materialsTotal ? String(materialsTotal) : "--";
  const mode = activeLibraryDoc ? "Lecture + General" : "General";
  return `LECTURE ${lecture} \u2014 MATERIALS ${materialsCount} \u2014 MODE ${mode}`;
}

function buildResponseContextBar() {
  const bar = document.createElement("div");
  bar.className = "response-context";
  bar.textContent = getContextLine();
  return bar;
}

function updateResponseContextBar() {
  document.querySelectorAll(".response-context").forEach(node => {
    node.textContent = getContextLine();
  });
}

function setResponseHighlightState(card, enabled, toggle) {
  if (!card) return;
  card.dataset.highlight = enabled ? "on" : "off";
  if (toggle) {
    toggle.classList.toggle("is-on", enabled);
    toggle.setAttribute("aria-checked", String(enabled));
  }
}

function buildResponseHeader(card) {
  const header = document.createElement("div");
  header.className = "response-header";
  const toolbarRow = document.createElement("div");
  toolbarRow.className = "response-toolbar-row";
  const highlightWrap = document.createElement("div");
  highlightWrap.className = "response-highlight-toggle response-toggle-row";
  const highlightToggle = document.createElement("button");
  highlightToggle.type = "button";
  highlightToggle.className = "toggle-switch is-on";
  highlightToggle.setAttribute("role", "switch");
  highlightToggle.setAttribute("aria-checked", "true");
  highlightToggle.setAttribute("aria-label", "Toggle highlighted key terms");
  highlightToggle.addEventListener("click", () => {
    const next = !highlightToggle.classList.contains("is-on");
    setResponseHighlightState(card, next, highlightToggle);
  });
  const highlightText = document.createElement("span");
  highlightText.textContent = "Highlighted key terms";
  highlightWrap.append(highlightToggle, highlightText);

  const actions = document.createElement("div");
  actions.className = "response-actions";
  const sources = document.createElement("button");
  sources.type = "button";
  sources.className = "response-sources";
  sources.dataset.role = "sources-count";
  sources.textContent = "0 sources";
  sources.addEventListener("click", () => openSourcesDrawer(card));

  const toolbar = document.createElement("div");
  toolbar.className = "response-toolbar";

  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "response-action";
  copyBtn.dataset.action = "copy";
  copyBtn.setAttribute("aria-label", "Copy answer");
  copyBtn.title = "Copy answer";
  copyBtn.innerHTML = `
    <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <rect x="9" y="9" width="11" height="11" rx="2"></rect>
      <rect x="4" y="4" width="11" height="11" rx="2"></rect>
    </svg>
  `;
  copyBtn.addEventListener("click", () => copyResponseText(card));

  const regenBtn = document.createElement("button");
  regenBtn.type = "button";
  regenBtn.className = "response-action";
  regenBtn.dataset.action = "regenerate";
  regenBtn.setAttribute("aria-label", "Regenerate answer");
  regenBtn.title = "Regenerate answer";
  regenBtn.innerHTML = `
    <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <path d="M21 12a9 9 0 1 1-2.6-6.4"></path>
      <polyline points="22 4 21 10 15 9"></polyline>
    </svg>
  `;
  regenBtn.addEventListener("click", () => regenerateResponse(card));

  const summaryBtn = document.createElement("button");
  summaryBtn.type = "button";
  summaryBtn.className = "response-action toggle-btn response-summary-toggle";
  summaryBtn.dataset.action = "toggle-summary";
  summaryBtn.setAttribute("aria-pressed", "false");
  summaryBtn.setAttribute("aria-label", "Toggle summary");
  summaryBtn.title = "Toggle summary";
  summaryBtn.innerHTML = `
    <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <line x1="4" y1="6" x2="20" y2="6"></line>
      <line x1="4" y1="12" x2="20" y2="12"></line>
      <line x1="4" y1="18" x2="14" y2="18"></line>
    </svg>
  `;
  summaryBtn.addEventListener("click", () => {
    const nextMode = card.dataset.viewMode === "summary" ? "full" : "summary";
    applyResponseView(card, nextMode, { persist: false });
  });

  const sourcesBtn = document.createElement("button");
  sourcesBtn.type = "button";
  sourcesBtn.className = "response-action";
  sourcesBtn.dataset.action = "view-sources";
  sourcesBtn.setAttribute("aria-label", "View sources");
  sourcesBtn.title = "View sources";
  sourcesBtn.innerHTML = `
    <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <path d="M4 4h16v14H4z"></path>
      <path d="M8 8h8"></path>
      <path d="M8 12h6"></path>
    </svg>
  `;
  sourcesBtn.addEventListener("click", () => openSourcesDrawer(card));

  toolbar.append(copyBtn, regenBtn, summaryBtn, sourcesBtn);
  actions.append(sources, toolbar);
  toolbarRow.append(actions);
  header.append(toolbarRow, highlightWrap);
  return header;
}

/**
 * Build a response card DOM node from assistant text and citations.
 *
 * @param text - Assistant markdown/text content.
 * @param citationMap - Map of citation ids to URL metadata.
 * @param opts - Rendering options (sources, msgId, section prefix, etc.).
 * @returns Response card element ready for insertion.
 * @remarks Side effects: mutates citation map and creates DOM nodes.
 */
function decorateAssistantResponse(text, citationMap = new Map(), opts = {}) {
  const rawAnswerText = typeof opts.answerText === "string" ? opts.answerText : (text || "");
  const inlineReferences = extractInlineReferences(rawAnswerText);
  const sanitizedText = sanitizeAssistantTextForSections(text);
  const preserveCitationMap = Boolean(opts?.preserveCitationMap);
  if (!preserveCitationMap) {
    enrichCitationMapFromText(text, citationMap);
    backfillCitationMapFromUrls(text, citationMap);
  }
  const inlineUrls = collectUrlsFromText(text);
  if (inlineUrls.length && !preserveCitationMap) {
    addInlineUrlsToCitationMap(citationMap, inlineUrls);
  }
  const responseCard = document.createElement("div");
  responseCard.className = "response-card";
  if (opts?.msgId) {
    responseCard.dataset.msgId = String(opts.msgId);
  }
  const rawPrefix = typeof opts?.sectionPrefix === "string" && opts.sectionPrefix.trim()
    ? opts.sectionPrefix.trim()
    : (opts?.msgId ? `msg-${opts.msgId}` : "");
  const sectionPrefix = normalizeAnchorPrefix(rawPrefix);
  if (sectionPrefix) {
    responseCard.dataset.sectionPrefix = sectionPrefix;
  }
  const header = buildResponseHeader(responseCard);
  const body = document.createElement("div");
  body.className = "response-body";
  const toc = document.createElement("div");
  toc.className = "response-toc";
  toc.hidden = true;
  const tocItems = document.createElement("div");
  tocItems.className = "response-toc-items";
  toc.append(tocItems);
  const content = document.createElement("div");
  content.className = "response-content";
  body.append(toc, content);
  responseCard.append(header, body);

  const trimmed = (sanitizedText || "").trim();
  const needsRich =
    /\n/.test(trimmed) ||
    /\|.+\|/.test(trimmed) ||
    /[-*+]\s+\w+/.test(trimmed) ||
    /```/.test(trimmed) ||
    /\*\*.+\*\*/.test(trimmed);

  if (trimmed && trimmed.length <= 400 && citationMap.size === 0 && inlineUrls.length === 0 && !needsRich) {
    const p = document.createElement("p");
    p.textContent = trimmed;
    content.appendChild(p);
  } else {
    const sections = parseSectionsFromText(trimmed);
    const usedIds = new Set();
    sections.forEach((sec, index) => {
      if (!sec || !sec.text) return;
      const blocks = sec.text
        .split(/\n{2,}/)
        .filter(Boolean)
        .filter(block => !isCitationsNoise(block))
        .filter(block => !isSourcesUnavailableBlock(block));
      if (!blocks.length) return;
      const section = document.createElement("section");
      section.className = "assistant-section response-section";
      const label = sec.label || "Response";
      const headingTag = sec.level === 3 ? "h3" : "h2";
      const heading = document.createElement(headingTag);
      heading.className = "section-heading";
      const normalizedLabel = normalizeHeadingCandidate(label).toLowerCase();
      const isResponseLabel = normalizedLabel === "response";
      if (isResponseLabel) {
        heading.classList.add("section-heading--response");
      }
      heading.textContent = `${pickEmoji(label)} ${label}`;
      const bodyInner = document.createElement("div");
      bodyInner.className = "section-body";
      blocks.forEach(block => {
        if (isReferenceBlock(block)) return;
        const htmlParts = splitHtmlTableBlocks(block);
        if (htmlParts) {
          const fragment = document.createDocumentFragment();
          let pendingCaption = "";
          let failed = false;
          htmlParts.forEach((part, idx) => {
            if (failed) return;
            if (part.type === "text") {
              const nextIsTable = htmlParts[idx + 1]?.type === "table";
              let text = part.text || "";
              if (nextIsTable) {
                const extracted = extractTrailingCaption(text);
                text = extracted.text;
                if (extracted.caption) pendingCaption = extracted.caption;
              }
              if (text && text.trim()) {
                fragment.appendChild(renderMarkdownContent(text));
              }
              return;
            }
            const tableRender = buildHtmlTable(part.html);
            if (!tableRender) {
              failed = true;
              return;
            }
            const captionText = pendingCaption || tableRender.caption;
            pendingCaption = "";
            if (captionText) {
              const captionNode = renderTableCaptionNode(captionText, { usedIds, idPrefix: sectionPrefix });
              if (captionNode) fragment.appendChild(captionNode);
            }
            fragment.appendChild(tableRender.node);
          });
          if (failed) {
            bodyInner.appendChild(renderMarkdownContent(block));
          } else {
            if (pendingCaption) {
              const captionNode = renderTableCaptionNode(pendingCaption, { usedIds, idPrefix: sectionPrefix });
              if (captionNode) fragment.appendChild(captionNode);
            }
            bodyInner.appendChild(fragment);
          }
          return;
        }
        const tableRender = buildMarkdownTable(block);
        if (tableRender) {
          if (tableRender.caption) {
            const captionNode = renderTableCaptionNode(tableRender.caption, { usedIds, idPrefix: sectionPrefix });
            if (captionNode) bodyInner.appendChild(captionNode);
          }
          bodyInner.appendChild(tableRender.node);
          if (tableRender.trailingText) {
            bodyInner.appendChild(renderMarkdownContent(tableRender.trailingText));
          }
        } else if (isTableLikeBlock(block)) {
          bodyInner.appendChild(renderTableFallback(block));
        } else {
          bodyInner.appendChild(renderMarkdownContent(block));
        }
      });
      resolveCitationLinks(bodyInner, citationMap);
      if (!sectionHasMeaningfulContent(bodyInner)) return;
      const sectionId = buildSectionId(label, index, usedIds, { prefix: sectionPrefix });
      heading.id = sectionId;
      section.dataset.sectionId = sectionId;
      section.append(heading, bodyInner);
      content.appendChild(section);
      renderMathIfReady(section);
    });
    if (!content.childNodes.length && trimmed) {
      const p = document.createElement("p");
      p.textContent = trimmed;
      content.appendChild(p);
    }
  }
  resolveCitationLinks(content, citationMap);
  responseState.set(responseCard, {
    citationMap,
    inlineUrls,
    sources: Array.isArray(opts.sources) ? opts.sources : [],
    inlineReferences,
    answerText: rawAnswerText,
  });
  return responseCard;
}

/**
 * Read or initialize response state for a response card.
 *
 * @param card - Response card element.
 * @returns Stored state object or null when card is missing.
 */
function getResponseState(card) {
  if (!card) return null;
  const existing = responseState.get(card) || {};
  if (!existing.citationMap) existing.citationMap = new Map();
  if (!Array.isArray(existing.sources)) existing.sources = [];
  if (!Array.isArray(existing.inlineUrls)) existing.inlineUrls = [];
  responseState.set(card, existing);
  return existing;
}

/**
 * Toggle all collapsible sections within a response card.
 *
 * @param card - Response card element.
 * @param open - Whether sections should be expanded.
 * @param opts - Optional expansion tracking flags.
 */
function setAllSectionsOpen(card, open, { markExpanded = false } = {}) {
  if (!card) return;
  if (markExpanded) {
    card.dataset.fullExpanded = "true";
  }
}

function ensureFullModeExpanded(card) {
  if (!card || card.dataset.fullExpanded === "true") return;
  setAllSectionsOpen(card, true, { markExpanded: true });
}

function updateToggleGroup(group, activeKey, datasetKey) {
  if (!group) return;
  group.querySelectorAll("button").forEach(btn => {
    const isActive = btn.dataset[datasetKey] === activeKey;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-pressed", String(isActive));
  });
}

/**
 * Apply the response view mode (summary/full) to a card.
 *
 * @param card - Response card element.
 * @param mode - View mode string.
 * @param opts - Persistence options for localStorage.
 */
function applyResponseView(card, mode, { persist = true } = {}) {
  if (!card) return;
  const normalized = normalizeViewMode(mode);
  card.dataset.viewMode = normalized;
  const summaryToggle = card.querySelector("[data-action=\"toggle-summary\"]");
  if (summaryToggle) {
    const isSummary = normalized === "summary";
    summaryToggle.setAttribute("aria-pressed", String(isSummary));
    summaryToggle.classList.toggle("is-active", isSummary);
    summaryToggle.title = isSummary ? "Show full answer" : "Show summary";
    summaryToggle.setAttribute("aria-label", isSummary ? "Show full answer" : "Show summary");
  }
  if (normalized === "summary") {
    buildResponseSummary(card);
  }
  if (persist) safeSetItem(RESPONSE_VIEW_KEY, normalized);
  ensureFullModeExpanded(card);
}

function applyCitationMode(card, mode, { persist = true } = {}) {
  if (!card) return;
  const normalized = normalizeCitationMode(mode);
  card.dataset.citeMode = normalized;
  updateToggleGroup(card.querySelector(".response-cite-toggle"), normalized, "citeMode");
  if (persist) safeSetItem(CITATION_MODE_KEY, normalized);
  ensureCitationEndnotes(card);
}

function installResponseCardCollapseToggle(card) {
  if (!card) return;
  const header = card.querySelector(".response-header");
  if (!header) return;
  const toolbar = header.querySelector(".response-toolbar") || header;
  if (toolbar.querySelector(".response-collapse-toggle")) return;

  let body = card.querySelector(".response-body");
  if (!body) {
    body = document.createElement("div");
    body.className = "response-body";
    const parent = header.parentElement || card;
    const siblings = Array.from(parent.children);
    const headerIndex = siblings.indexOf(header);
    const toMove = siblings.slice(headerIndex + 1);
    toMove.forEach(node => body.appendChild(node));
    parent.appendChild(body);
  }

  const bodyId = ensureId(body, "response-body");
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "response-collapse-toggle";
  btn.setAttribute("aria-controls", bodyId);
  const chevron = document.createElement("span");
  chevron.className = "chevron";
  btn.appendChild(chevron);

  const setCollapsed = (collapsed) => {
    card.dataset.collapsed = collapsed ? "true" : "false";
    if (body.hasAttribute("hidden")) body.removeAttribute("hidden");
    body.setAttribute("aria-hidden", collapsed ? "true" : "false");
    btn.setAttribute("aria-expanded", collapsed ? "false" : "true");
    btn.setAttribute("aria-label", collapsed ? "Expand answer" : "Collapse answer");
    btn.title = collapsed ? "Expand answer" : "Collapse answer";
    if (collapsed) {
      const active = document.activeElement;
      if (active && body.contains(active)) {
        btn.focus();
      }
    }
  };

  setCollapsed(false);
  btn.addEventListener("click", () => {
    const collapsed = card.dataset.collapsed === "true";
    setCollapsed(!collapsed);
  });

  toolbar.appendChild(btn);
}

function cleanSummaryHeading(text) {
  return (text || "").replace(/^[^\w]+/, "").trim();
}

function trimSummaryText(text, { sentences = 2, maxChars = 320 } = {}) {
  const cleaned = (text || "").replace(/\s+/g, " ").trim();
  if (!cleaned) return "";
  const parts = cleaned.match(/[^.!?]+[.!?]+|[^.!?]+$/g);
  let clipped = parts ? parts.slice(0, sentences).join(" ").trim() : cleaned;
  if (clipped.length > maxChars) {
    clipped = `${clipped.slice(0, Math.max(0, maxChars - 1))}…`;
  }
  return clipped;
}

function buildResponseSummary(card) {
  if (!card) return;
  const body = card.querySelector(".response-body");
  const content = card.querySelector(".response-content");
  if (!body || !content) return;
  if (body.querySelector(".response-summary")) return;

  const summary = document.createElement("div");
  summary.className = "response-summary";
  const summaryHeader = document.createElement("div");
  summaryHeader.className = "response-summary-header";
  summaryHeader.textContent = "Summary";
  summary.appendChild(summaryHeader);

  const firstSection = content.querySelector(".response-section");
  const firstHeading = firstSection?.querySelector(".section-heading");
  const headingText = cleanSummaryHeading(firstHeading?.textContent || "");
  const leadRoot = firstSection?.querySelector(".section-body") || content;
  const leadCandidate = leadRoot.querySelector("p, li, blockquote");
  const leadText = trimSummaryText(leadCandidate?.textContent || "");
  if (headingText || leadText) {
    const lead = document.createElement("div");
    lead.className = "response-summary-lead";
    if (headingText) {
      const heading = document.createElement("div");
      heading.className = "response-summary-heading";
      heading.textContent = headingText;
      lead.appendChild(heading);
    }
    if (leadText) {
      const p = document.createElement("p");
      p.textContent = leadText;
      lead.appendChild(p);
    }
    summary.appendChild(lead);
  }

  const list = document.createElement("ul");
  list.className = "response-summary-list";
  const bullets = Array.from(content.querySelectorAll("li"))
    .filter(li => (li.textContent || "").trim())
    .slice(0, RESPONSE_SUMMARY_BULLET_LIMIT);
  bullets.forEach(li => {
    const clone = li.cloneNode(true);
    clone.querySelectorAll("ul,ol").forEach(nested => nested.remove());
    list.appendChild(clone);
  });
  if (list.childNodes.length) {
    summary.appendChild(list);
  }

  const table = content.querySelector(".owen-table-card, .AnswerTable, table");
  if (table) {
    const note = document.createElement("div");
    note.className = "response-summary-note";
    note.textContent = "Table available in Full view.";
    summary.appendChild(note);
  }

  if (!summary.querySelector(".response-summary-lead") && !list.childNodes.length) {
    const fallbackText = trimSummaryText(content.textContent || "", {
      sentences: 2,
      maxChars: SECTION_PREVIEW_CHAR_LIMIT,
    });
    if (fallbackText) {
      const p = document.createElement("p");
      p.className = "response-summary-fallback";
      p.textContent = fallbackText;
      summary.appendChild(p);
    }
  }

  body.insertBefore(summary, content);
  const state = getResponseState(card);
  if (state?.normalizedSources) {
    registerCitationChips(card, state.normalizedSources);
  }
  card.querySelectorAll(".citation-endnotes").forEach(node => node.remove());
  card.dataset.citePrepared = "false";
  ensureCitationEndnotes(card);
}

function getScrollMarginTop(node) {
  if (!node || !window.getComputedStyle) return 0;
  const marginTop = getComputedStyle(node).scrollMarginTop || "0";
  const parsed = parseFloat(marginTop);
  return Number.isFinite(parsed) ? parsed : 0;
}

function scrollResponseTarget(card, target, opts = {}) {
  if (!card || !target) return;
  const behavior = opts.behavior === "auto" ? "auto" : "smooth";
  const updateHash = Boolean(opts.updateHash);
  if (updateHash && target.id) {
    updateUrlHash(target.id);
  }
  const body = card.querySelector(".response-body");
  const anchorOffset = getScrollMarginTop(target);
  if (body && body.scrollHeight > body.clientHeight) {
    const bodyRect = body.getBoundingClientRect();
    const targetRect = target.getBoundingClientRect();
    const offset = targetRect.top - bodyRect.top + body.scrollTop - anchorOffset;
    const top = Math.max(0, Math.round(offset));
    if (typeof body.scrollTo === "function") {
      body.scrollTo({ top, behavior });
    } else {
      body.scrollTop = top;
    }
    return;
  }
  if (typeof target.scrollIntoView === "function") {
    target.scrollIntoView({ behavior, block: "start" });
  }
}

let lastHashScrollId = "";

function getHashTargetId() {
  const raw = window.location.hash || "";
  if (!raw || raw === "#") return "";
  try {
    return decodeURIComponent(raw.slice(1));
  } catch {
    return raw.slice(1);
  }
}

function updateUrlHash(id, { replace = true } = {}) {
  if (!id) return;
  const encoded = `#${encodeURIComponent(id)}`;
  if (window.location.hash === encoded) {
    lastHashScrollId = id;
    return;
  }
  if (replace && history?.replaceState) {
    history.replaceState(null, "", encoded);
  } else {
    window.location.hash = encoded;
  }
  lastHashScrollId = id;
}

function scrollToHash({ behavior = "smooth" } = {}) {
  const id = getHashTargetId();
  if (!id || id === lastHashScrollId) return false;
  const target = document.getElementById(id);
  if (!target) return false;
  const card = target.closest?.(".response-card");
  if (card) {
    scrollResponseTarget(card, target, { behavior });
    setActiveTocItem(card, id);
  } else if (typeof target.scrollIntoView === "function") {
    target.scrollIntoView({ behavior, block: "start" });
  }
  lastHashScrollId = id;
  return true;
}

function buildResponseToc(card) {
  const toc = card?.querySelector(".response-toc");
  const tocItems = toc?.querySelector(".response-toc-items");
  if (!toc || !tocItems) return [];
  tocItems.replaceChildren();
  const sections = Array.from(card.querySelectorAll(".response-section"));
  const sectionHeadings = sections
    .map(section => section.querySelector(".section-heading"))
    .filter(Boolean);
  const usedIds = new Set(sectionHeadings.map(node => node.id).filter(Boolean));
  const sectionPrefix = normalizeAnchorPrefix(card?.dataset?.sectionPrefix || "");
  const items = [];

  const addItem = (targetNode, text) => {
    if (!text) return;
    const id = targetNode.id || buildSectionId(text, items.length, usedIds, { prefix: sectionPrefix });
    targetNode.id = id;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "toc-item";
    btn.dataset.targetId = id;
    btn.textContent = text.replace(/^[^\w]+/, "").trim();
    btn.addEventListener("click", () => {
      const target = document.getElementById(id);
      if (!target) return;
      scrollResponseTarget(card, target, { updateHash: true });
      setActiveTocItem(card, id);
    });
    tocItems.appendChild(btn);
    items.push({ id, node: targetNode, button: btn });
  };

  sectionHeadings.forEach(node => addItem(node, (node.textContent || "").trim()));
  toc.hidden = items.length < 2;
  if (toc.hidden) {
    card?.style?.removeProperty("--response-anchor-offset");
  } else if (card) {
    const height = toc.getBoundingClientRect().height || 0;
    const anchorOffset = Math.max(56, Math.round(height) + 12);
    card.style.setProperty("--response-anchor-offset", `${anchorOffset}px`);
  }
  return items;
}

function setActiveTocItem(card, id) {
  const toc = card?.querySelector(".response-toc");
  if (!toc) return;
  let changed = false;
  toc.querySelectorAll(".toc-item").forEach(btn => {
    const isActive = btn.dataset.targetId === id;
    if (btn.classList.contains("is-active") !== isActive) {
      btn.classList.toggle("is-active", isActive);
      if (isActive) {
        btn.setAttribute("aria-current", "true");
      } else {
        btn.removeAttribute("aria-current");
      }
      changed = true;
    }
  });
  if (changed) {
    const active = toc.querySelector(".toc-item.is-active");
    const scroller = toc.querySelector(".response-toc-items");
    if (active && scroller && scroller.scrollWidth > scroller.clientWidth && typeof active.scrollIntoView === "function") {
      active.scrollIntoView({ block: "nearest", inline: "center" });
    }
  }
}

function setupTocObserver(card, items) {
  if (!items.length) return;
  const body = card?.querySelector(".response-body") || null;
  const scrollRoot = body && body.scrollHeight > body.clientHeight ? body : (chatLog || null);
  if (!("IntersectionObserver" in window)) {
    const root = scrollRoot || body || chatLog;
    if (!root) return;
    const handler = () => {
      const bodyRect = root.getBoundingClientRect();
      let best = null;
      let bestOffset = Infinity;
      items.forEach(item => {
        const rect = item.node.getBoundingClientRect();
        const offset = rect.top - bodyRect.top;
        if (offset >= -4 && offset < bestOffset) {
          bestOffset = offset;
          best = item;
        }
      });
      if (!best) {
        const above = items.filter(item => item.node.getBoundingClientRect().top < bodyRect.top);
        if (above.length) best = above[above.length - 1];
      }
      if (best?.id) setActiveTocItem(card, best.id);
    };
    root.addEventListener("scroll", handler, { passive: true });
    handler();
    return;
  }
  const observer = new IntersectionObserver((entries) => {
    const visible = entries
      .filter(entry => entry.isIntersecting)
      .sort((a, b) => b.intersectionRatio - a.intersectionRatio);
    if (visible[0]?.target?.id) {
      setActiveTocItem(card, visible[0].target.id);
    }
  }, {
    root: scrollRoot || null,
    rootMargin: "0px 0px -70% 0px",
    threshold: [0.1, 0.25, 0.5],
  });
  items.forEach(item => observer.observe(item.node));
  const state = getResponseState(card);
  if (state) state.tocObserver = observer;
}

function registerCitationChips(card, sources) {
  if (!card) return;
  const sourceMap = new Map();
  (sources || []).forEach(src => {
    if (!src || !src.id) return;
    sourceMap.set(String(src.id), src);
    sourceMap.set(Number(src.id), src);
  });
  const hasSources = sourceMap.size > 0;
  card.querySelectorAll(".citation-chip").forEach(chip => {
    let id = chip.dataset.citeId || "";
    if (!id) {
      id = String((chip.textContent || "").replace(/[^\d]/g, ""));
      if (id) chip.dataset.citeId = id;
    }
    if (!chip.dataset.citeInline) chip.dataset.citeInline = "true";
    if (!hasSources) {
      chip.classList.add("cite--disabled");
      chip.dataset.citeDisabled = "true";
      chip.setAttribute("aria-disabled", "true");
      chip.setAttribute("tabindex", "-1");
      chip.title = "No sources returned for this response.";
      if (chip.hasAttribute("href")) {
        chip.removeAttribute("href");
        chip.removeAttribute("target");
        chip.removeAttribute("rel");
      }
      return;
    }
    chip.classList.remove("cite--disabled");
    chip.dataset.citeDisabled = "false";
    chip.removeAttribute("aria-disabled");
    if (chip.getAttribute("tabindex") === "-1") chip.removeAttribute("tabindex");
    const info = sourceMap.get(id) || sourceMap.get(Number(id)) || null;
    const label = info?.title || info?.domain || info?.url || `Source ${id || ""}`.trim();
    if (id) {
      chip.setAttribute("aria-label", `Open source ${id}: ${label}`);
    }
    chip.title = label;
    bindCitationHover(chip);
  });
  const state = getResponseState(card);
  if (state) state.sourceMap = sourceMap;
}

function bindCitationHover(chip) {
  if (!chip || chip.dataset.tooltipBound === "true") return;
  if (chip.dataset.citeDisabled === "true" || chip.getAttribute("aria-disabled") === "true") return;
  chip.dataset.tooltipBound = "true";
  chip.addEventListener("pointerenter", () => {
    tooltipHoveringChip = true;
    showCitationTooltip(chip);
  });
  chip.addEventListener("pointerleave", () => {
    tooltipHoveringChip = false;
    scheduleTooltipHide();
  });
}

function ensureCitationEndnotes(card) {
  if (!card || card.dataset.citePrepared === "true") return;
  const state = getResponseState(card);
  const sourceMap = state?.sourceMap || new Map();
  const content = card.querySelector(".response-content");
  const summary = card.querySelector(".response-summary");
  if (!content && !summary) return;
  const blocks = [
    ...(content ? Array.from(content.querySelectorAll("p, li, td, th")) : []),
    ...(summary ? Array.from(summary.querySelectorAll("p, li, td, th")) : []),
  ];
  blocks.forEach(block => {
    const chips = Array.from(block.querySelectorAll(".citation-chip[data-cite-id]"));
    if (!chips.length) return;
    const ids = Array.from(new Set(chips.map(chip => chip.dataset.citeId).filter(Boolean)));
    if (!ids.length) return;
    const endnotes = document.createElement("span");
    endnotes.className = "citation-endnotes";
    ids.forEach(id => {
      const info = sourceMap.get(id) || sourceMap.get(Number(id));
      if (!info) return;
      const chip = buildCitationPill(info);
      if (chip?.dataset) {
        chip.dataset.citeInline = "false";
      }
      const anchor = chip?.querySelector ? chip.querySelector(".citation-chip") : null;
      if (anchor) {
        anchor.dataset.citeId = String(info.id);
        anchor.dataset.citeInline = "false";
      }
      endnotes.appendChild(chip);
    });
    block.appendChild(endnotes);
  });
  card.dataset.citePrepared = "true";
}

function updateSourcesCount(card, { verifiedCount = 0, inlineCount = 0 } = {}) {
  const label = card.querySelector("[data-role=\"sources-count\"]");
  if (!label) return;
  const hasVerified = Number.isFinite(verifiedCount) && verifiedCount > 0;
  const hasInline = Number.isFinite(inlineCount) && inlineCount > 0;

  if (hasVerified) {
    label.textContent = `${verifiedCount} source${verifiedCount === 1 ? "" : "s"}`;
  } else if (hasInline) {
    label.textContent = `${inlineCount} source${inlineCount === 1 ? "" : "s"}`;
  } else {
    label.textContent = "0 sources";
  }
  label.hidden = false;

  const viewBtn = card.querySelector("[data-action=\"view-sources\"]");
  if (viewBtn) {
    const enabled = hasVerified || hasInline;
    viewBtn.disabled = !enabled;
    viewBtn.title = enabled ? "" : "No sources returned for this response.";
    viewBtn.setAttribute("aria-disabled", String(!enabled));
  }
  if (label instanceof HTMLButtonElement) {
    label.disabled = !(hasVerified || hasInline);
    label.setAttribute("aria-disabled", String(!(hasVerified || hasInline)));
  }
}

/**
 * Finalize a response card after content is inserted.
 *
 * @param card - Response card element.
 * @remarks Side effects: registers citation pills, TOC, and collapse handlers.
 */
function finalizeResponseCard(card) {
  if (!card || card.dataset.prepared === "true") return;
  const state = getResponseState(card);
  const sources = normalizeCitedSources(state?.sources);
  const inlineReferences = state?.inlineReferences || extractInlineReferences(state?.answerText || "");
  if (state) {
    state.normalizedSources = sources;
    state.inlineReferences = inlineReferences;
  }
  registerCitationChips(card, sources);
  updateSourcesCount(card, {
    verifiedCount: sources.length,
    inlineCount: inlineReferences.items.length,
  });
  const tocItems = buildResponseToc(card);
  setupTocObserver(card, tocItems);
  scrollToHash({ behavior: "auto" });
  ensureCitationEndnotes(card);
  applyResponseView(card, "full", { persist: false });
  applyCitationMode(card, "inline", { persist: false });
  installResponseCardCollapseToggle(card);
  const highlightToggle = card.querySelector(".response-highlight-toggle");
  const hasHighlights = Boolean(card.querySelector(".response-content strong, .response-summary strong"));
  if (highlightToggle) {
    highlightToggle.hidden = !hasHighlights;
    setResponseHighlightState(card, hasHighlights, highlightToggle.querySelector(".toggle-switch"));
  }
  card.dataset.prepared = "true";
}

async function copyResponseText(card) {
  if (!card) return;
  const state = getResponseState(card);
  const text = (state?.answerText || "").trim();
  if (!text) return;
  const btn = card.querySelector("[data-action=\"copy\"]");
  const originalLabel = btn?.getAttribute("aria-label") || "Copy answer";
  const originalTitle = btn?.title || "Copy answer";
  const showCopied = () => {
    if (!btn) return;
    btn.setAttribute("data-copied", "true");
    btn.setAttribute("aria-label", "Copied");
    btn.title = "Copied";
    setTimeout(() => {
      btn.removeAttribute("data-copied");
      btn.setAttribute("aria-label", originalLabel);
      btn.title = originalTitle;
    }, 1200);
  };
  try {
    await navigator.clipboard.writeText(text);
    showCopied();
  } catch {
    const fallback = document.createElement("textarea");
    fallback.value = text;
    fallback.style.position = "fixed";
    fallback.style.opacity = "0";
    document.body.appendChild(fallback);
    fallback.focus();
    fallback.select();
    try {
      document.execCommand("copy");
      showCopied();
    } finally {
      fallback.remove();
    }
  }
}

function regenerateResponse(card) {
  if (!card) return;
  const msgId = card.dataset.msgId || card.closest(".bubble")?.dataset?.msgId || "";
  let prompt = "";
  if (msgId) {
    const idx = conversation.findIndex(entry => entry?.id === msgId);
    if (idx > 0) {
      for (let i = idx - 1; i >= 0; i -= 1) {
        const entry = conversation[i];
        if (entry?.role === "user" && typeof entry.text === "string") {
          prompt = entry.text;
          break;
        }
      }
    }
  }
  if (!prompt) {
    const lastUser = [...conversation].reverse().find(entry => entry?.role === "user" && entry.text);
    prompt = lastUser?.text || "";
  }
  if (prompt) runChat(prompt);
}

function buildFaviconBadge(domain) {
  const badge = document.createElement("span");
  badge.className = "source-favicon";
  const letter = (domain || "").replace(/^www\./i, "").charAt(0).toUpperCase();
  badge.textContent = letter || "•";
  badge.setAttribute("aria-hidden", "true");
  return badge;
}

function renderSourcesDrawer(card) {
  if (!sourcesDrawer || !sourcesList) return;
  const state = getResponseState(card);
  const sources = state?.normalizedSources || [];
  const inlineReferences = state?.inlineReferences?.items || [];
  const filter = (sourcesSearch?.value || "").trim().toLowerCase();
  sourcesList.replaceChildren();
  const sourceMatches = sources.filter(src => {
    if (!filter) return true;
    const haystack = `${src.title || ""} ${src.domain || ""} ${src.url || ""}`.toLowerCase();
    return haystack.includes(filter);
  });
  const referenceMatches = inlineReferences.filter(ref => {
    if (!filter) return true;
    return ref.toLowerCase().includes(filter);
  });
  const hasSources = sources.length > 0;
  const hasReferences = inlineReferences.length > 0;

  if (!hasSources && !hasReferences) {
    const empty = document.createElement("div");
    empty.className = "sources-empty";
    empty.textContent = filter ? "No matching sources." : "No sources available.";
    sourcesList.appendChild(empty);
  } else {
    if (hasSources) {
      if (!sourceMatches.length) {
        const empty = document.createElement("div");
        empty.className = "sources-empty";
        empty.textContent = "No matching sources.";
        sourcesList.appendChild(empty);
      } else {
        sourceMatches.forEach(src => {
          const row = document.createElement("div");
          row.className = "source-row";
          row.dataset.sourceId = String(src.id);
          row.setAttribute("tabindex", "0");
          row.setAttribute("aria-label", `Source ${src.id}: ${src.title || src.domain || src.url || "Source"}`);
          const idx = document.createElement("span");
          idx.className = "source-index";
          idx.textContent = `[${src.id}]`;
          const body = document.createElement("div");
          body.className = "source-body";
          const title = document.createElement("div");
          title.className = "source-title";
          title.textContent = src.title || src.domain || src.url || "Source";
          const meta = document.createElement("div");
          meta.className = "source-meta";
          const domain = src.domain || getCitationDomain(src.url || "");
          const domainText = document.createElement("span");
          domainText.className = "source-domain";
          domainText.textContent = domain || "Source";
          meta.append(buildFaviconBadge(domain), domainText);
          const snippet = document.createElement("div");
          snippet.className = "source-snippet";
          snippet.textContent = src.snippet || "No snippet available.";
          const actions = document.createElement("div");
          actions.className = "source-actions";
          if (src.url) {
            const open = document.createElement("a");
            open.className = "sources-link";
            open.href = src.url;
            open.target = "_blank";
            open.rel = "noreferrer";
            open.textContent = "Open link";
            actions.appendChild(open);
          }
          body.append(title, meta, snippet, actions);
          row.append(idx, body);
          row.addEventListener("click", () => {
            setActiveSource(card, String(src.id));
          });
          row.addEventListener("keydown", (event) => {
            if (event.key === "Enter" || event.key === " ") {
              event.preventDefault();
              setActiveSource(card, String(src.id));
            }
          });
          sourcesList.appendChild(row);
        });
      }
    }

    if (hasReferences) {
      const section = document.createElement("div");
      section.className = "sources-drawer__section";
      if (hasSources) section.classList.add("sources-drawer__section--divider");
      const heading = document.createElement("div");
      heading.className = "sources-drawer__section-title";
      heading.textContent = "References from answer (unverified)";
      heading.title = "Extracted from the answer text; not retrieval-verified.";
      section.appendChild(heading);

      if (!referenceMatches.length) {
        const empty = document.createElement("div");
        empty.className = "sources-empty";
        empty.textContent = "No matching references.";
        section.appendChild(empty);
      } else {
        const list = document.createElement("ol");
        list.className = "references-list";
        referenceMatches.forEach(item => {
          const li = document.createElement("li");
          li.textContent = item;
          list.appendChild(li);
        });
        section.appendChild(list);
      }
      sourcesList.appendChild(section);
    }
  }
  const title = sourcesDrawer.querySelector("[data-role=\"sources-title\"]");
  if (title) {
    if (hasSources) {
      title.textContent = `Sources (${sources.length})`;
    } else if (hasReferences) {
      title.textContent = `References (${inlineReferences.length})`;
    } else {
      title.textContent = "Sources (0)";
    }
  }
  highlightSourceRow(activeSourceId);
}

function highlightSourceRow(sourceId) {
  if (!sourcesList) return;
  sourcesList.querySelectorAll(".source-row").forEach(row => {
    row.classList.toggle("is-active", row.dataset.sourceId === sourceId);
  });
}

function highlightCitationChips(card, sourceId) {
  if (!card) return;
  card.querySelectorAll(".citation-chip.is-highlighted").forEach(chip => {
    chip.classList.remove("is-highlighted");
  });
  if (!sourceId) return;
  card.querySelectorAll(`.citation-chip[data-cite-id="${sourceId}"]`).forEach(chip => {
    chip.classList.add("is-highlighted");
  });
}

function scrollToCitation(card, sourceId) {
  if (!card || !sourceId) return;
  const target = card.querySelector(`.citation-chip[data-cite-id="${sourceId}"]`);
  if (!target) return;
  target.scrollIntoView({ behavior: "smooth", block: "center" });
}

function setActiveSource(card, sourceId) {
  activeSourceId = sourceId;
  highlightSourceRow(sourceId);
  highlightCitationChips(card, sourceId);
  scrollToCitation(card, sourceId);
}

function openSourcesDrawer(card, sourceId) {
  if (!sourcesDrawer || !sourcesList) return;
  const targetCard = card || activeSourcesCard;
  if (!targetCard) return;
  activeSourcesCard = targetCard;
  if (targetCard.dataset.prepared !== "true") finalizeResponseCard(targetCard);
  renderSourcesDrawer(targetCard);
  sourcesDrawer.classList.add("is-open");
  sourcesDrawer.setAttribute("aria-hidden", "false");
  if (sourcesOverlay) sourcesOverlay.hidden = false;
  document.body.classList.add("sources-open");
  if (sourceId) setActiveSource(targetCard, String(sourceId));
  if (sourcesSearch) sourcesSearch.focus();
}

function closeSourcesDrawer() {
  if (!sourcesDrawer) return;
  sourcesDrawer.classList.remove("is-open");
  sourcesDrawer.setAttribute("aria-hidden", "true");
  if (sourcesOverlay) sourcesOverlay.hidden = true;
  document.body.classList.remove("sources-open");
  activeSourceId = null;
  highlightSourceRow(null);
  if (activeSourcesCard) highlightCitationChips(activeSourcesCard, null);
}

function getFocusableElements(container) {
  return Array.from(container.querySelectorAll("button, [href], input, select, textarea, [tabindex]:not([tabindex=\"-1\"])"))
    .filter(el => !el.disabled && !el.hasAttribute("disabled") && el.getClientRects().length > 0);
}

function handleDrawerKeydown(event) {
  if (!sourcesDrawer || !sourcesDrawer.classList.contains("is-open")) return;
  if (event.key === "Escape") {
    event.preventDefault();
    closeSourcesDrawer();
    return;
  }
  if (event.key !== "Tab") return;
  const focusable = getFocusableElements(sourcesDrawer);
  if (!focusable.length) return;
  const first = focusable[0];
  const last = focusable[focusable.length - 1];
  if (event.shiftKey) {
    if (document.activeElement === first || !sourcesDrawer.contains(document.activeElement)) {
      event.preventDefault();
      last.focus();
    }
    return;
  }
  if (document.activeElement === last) {
    event.preventDefault();
    first.focus();
  }
}

const CITATION_TOOLTIP_ID = "citation-tooltip";
const TOOLTIP_HIDE_DELAY_MS = 120;
let citationTooltip = null;
let tooltipHideTimer = null;
let tooltipAnchor = null;
let openCitationId = null;
let tooltipHoveringChip = false;
let tooltipHoveringCard = false;

function ensureCitationTooltip() {
  if (citationTooltip) return citationTooltip;
  citationTooltip = document.createElement("div");
  citationTooltip.className = "citation-tooltip";
  citationTooltip.setAttribute("role", "tooltip");
  citationTooltip.setAttribute("aria-hidden", "true");
  citationTooltip.id = CITATION_TOOLTIP_ID;
  citationTooltip.hidden = true;
  citationTooltip.addEventListener("pointerenter", () => {
    tooltipHoveringCard = true;
    if (tooltipHideTimer) {
      clearTimeout(tooltipHideTimer);
      tooltipHideTimer = null;
    }
  });
  citationTooltip.addEventListener("pointerleave", () => {
    tooltipHoveringCard = false;
    scheduleTooltipHide();
  });
  document.body.appendChild(citationTooltip);
  return citationTooltip;
}

function getCitationInfoFromChip(chip) {
  if (!chip) return null;
  const card = chip.closest(".response-card");
  const state = getResponseState(card);
  const id = chip.dataset.citeId || String((chip.textContent || "").replace(/[^\d]/g, ""));
  if (!id) return null;
  let info = state?.sourceMap?.get(id) || state?.sourceMap?.get(Number(id)) || null;
  if (!info && state?.citationMap instanceof Map) {
    const raw = state.citationMap.get(id) || state.citationMap.get(Number(id));
    const url = typeof raw === "string" ? raw : raw?.url;
    if (url) {
      info = {
        id,
        url,
        title: raw?.title || "",
        domain: raw?.domain || getCitationDomain(url),
        snippet: raw?.snippet || "",
      };
    }
  }
  return { card, id, info };
}

function scheduleTooltipHide() {
  if (tooltipHideTimer) clearTimeout(tooltipHideTimer);
  tooltipHideTimer = setTimeout(() => {
    if (shouldKeepTooltipOpen()) return;
    hideCitationTooltip();
  }, TOOLTIP_HIDE_DELAY_MS);
}

function shouldKeepTooltipOpen() {
  if (!openCitationId) return false;
  const active = document.activeElement instanceof Element ? document.activeElement : null;
  const hasFocus = Boolean(
    active && ((tooltipAnchor && tooltipAnchor.contains(active)) || (citationTooltip && citationTooltip.contains(active))),
  );
  return tooltipHoveringChip || tooltipHoveringCard || hasFocus;
}

function hideCitationTooltip() {
  if (tooltipHideTimer) {
    clearTimeout(tooltipHideTimer);
    tooltipHideTimer = null;
  }
  if (!citationTooltip) return;
  citationTooltip.hidden = true;
  citationTooltip.setAttribute("aria-hidden", "true");
  if (tooltipAnchor) {
    tooltipAnchor.setAttribute("aria-expanded", "false");
  }
  tooltipAnchor = null;
  openCitationId = null;
  tooltipHoveringChip = false;
  tooltipHoveringCard = false;
}

function positionTooltip(anchor, tooltip) {
  const rect = anchor.getBoundingClientRect();
  const tipRect = tooltip.getBoundingClientRect();
  let top = rect.bottom + 10;
  if (top + tipRect.height > window.innerHeight) {
    top = rect.top - tipRect.height - 10;
  }
  let left = rect.left + rect.width / 2 - tipRect.width / 2;
  left = Math.max(12, Math.min(left, window.innerWidth - tipRect.width - 12));
  tooltip.style.top = `${Math.max(12, top)}px`;
  tooltip.style.left = `${left}px`;
}

function showCitationTooltip(chip) {
  if (chip?.dataset?.citeDisabled === "true" || chip?.getAttribute?.("aria-disabled") === "true") {
    hideCitationTooltip();
    return;
  }
  const payload = getCitationInfoFromChip(chip);
  if (!payload) {
    hideCitationTooltip();
    return;
  }
  const { card, id, info } = payload;
  if (!info?.url) {
    hideCitationTooltip();
    return;
  }
  const tooltip = ensureCitationTooltip();
  if (tooltipHideTimer) {
    clearTimeout(tooltipHideTimer);
    tooltipHideTimer = null;
  }
  if (tooltipAnchor && tooltipAnchor !== chip) {
    tooltipAnchor.setAttribute("aria-expanded", "false");
  }
  tooltipAnchor = chip;
  openCitationId = id;
  tooltipAnchor.setAttribute("aria-expanded", "true");
  tooltipAnchor.setAttribute("aria-controls", CITATION_TOOLTIP_ID);
  tooltip.replaceChildren();

  const title = document.createElement("div");
  title.className = "citation-tooltip-title";
  title.textContent = info.title || info.domain || info.url;
  tooltip.appendChild(title);

  const meta = document.createElement("div");
  meta.className = "citation-tooltip-meta";
  meta.append(buildFaviconBadge(info.domain || getCitationDomain(info.url)));
  const domain = document.createElement("span");
  domain.textContent = info.domain || getCitationDomain(info.url);
  meta.appendChild(domain);
  tooltip.appendChild(meta);

  if (info.snippet) {
    const snippet = document.createElement("div");
    snippet.className = "citation-tooltip-snippet";
    snippet.textContent = info.snippet.slice(0, 220);
    tooltip.appendChild(snippet);
  }

  const actions = document.createElement("div");
  actions.className = "citation-tooltip-actions";
  const openBtn = document.createElement("button");
  openBtn.type = "button";
  openBtn.textContent = "Open";
  openBtn.addEventListener("click", () => {
    window.open(info.url, "_blank", "noopener,noreferrer");
  });
  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.textContent = "Copy link";
  copyBtn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(info.url);
      copyBtn.textContent = "Copied";
      setTimeout(() => { copyBtn.textContent = "Copy link"; }, 1200);
    } catch {}
  });
  const showBtn = document.createElement("button");
  showBtn.type = "button";
  showBtn.textContent = "Show in Sources";
  showBtn.addEventListener("click", () => {
    openSourcesDrawer(card, id);
  });
  actions.append(openBtn, copyBtn, showBtn);
  tooltip.appendChild(actions);

  tooltip.hidden = false;
  tooltip.setAttribute("aria-hidden", "false");
  requestAnimationFrame(() => positionTooltip(chip, tooltip));
}

function showThinkingBubble() {
  return createScaffoldBubble();
}

function createScaffoldBubble(msgId) {
  const bubble = appendChatMessage("assistant", "", { track: false, msgId });
  bubble.classList.add("thinking", "scaffold");
  const textBlock = ensureBubbleTextNode(bubble, { reset: true });
  const status = document.createElement("div");
  status.className = "bubble-status";
  const statusText = document.createElement("span");
  statusText.className = "bubble-status-text";
  statusText.textContent = "O.W.E.N. Is Thinking";
  statusText.classList.add("thinking-text-strong");
  const interruptBtn = document.createElement("button");
  interruptBtn.type = "button";
  interruptBtn.className = "interrupt-btn";
  interruptBtn.innerHTML = `${INTERRUPT_ICON}<span>Interrupt</span>`;
  interruptBtn.setAttribute("aria-label", "Interrupt response");
  interruptBtn.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    interruptActiveThinking();
  });
  status.append(statusText, interruptBtn);
  const loadingBar = document.createElement("div");
  loadingBar.className = "loading-bar";
  const fill = document.createElement("div");
  fill.className = "loading-bar__fill";
  loadingBar.appendChild(fill);
  textBlock.append(status, loadingBar);
  return bubble;
}

function setBubbleStatus(bubble, text) {
  if (!bubble) return;
  const status = ensureBubbleStatusNode(bubble);
  const label = status?.querySelector(".bubble-status-text");
  if (label) label.textContent = text;
}

function getTypewriterController(bubble) {
  if (!bubble) return null;
  const id = bubble.dataset.msgId || generateMessageId();
  bubble.dataset.msgId = id;
  if (typewriterControllers.has(id)) return typewriterControllers.get(id);
  return null;
}

function startTypewriter(bubble, initialText = "", { onDone, completeOnSkip = false, preserveStatus = false } = {}) {
  if (!bubble) return null;
  const id = bubble.dataset.msgId || generateMessageId();
  bubble.dataset.msgId = id;
  const existingState = messageState.get(id) || { streamText: "", lastAnswer: "" };
  messageState.set(id, { ...existingState, streamText: existingState.streamText || initialText, lastAnswer: existingState.lastAnswer || initialText });
  const existing = typewriterControllers.get(id);
  if (existing) {
    existing.fullText = initialText;
    existing.completeOnSkip = completeOnSkip;
    existing.onDone = onDone || existing.onDone;
    return existing;
  }
  let statusNode = null;
  let loadingBar = null;
  if (preserveStatus) {
    statusNode = bubble.querySelector(".bubble-status");
    loadingBar = bubble.querySelector(".loading-bar");
  }
  const textBlock = ensureBubbleTextNode(bubble, { reset: true });
  if (preserveStatus) {
    if (statusNode) textBlock.appendChild(statusNode);
    if (loadingBar) textBlock.appendChild(loadingBar);
  }
  const plain = document.createElement("div");
  plain.className = "typewriter-plain";
  plain.textContent = "";
  textBlock.appendChild(plain);
  const controller = {
    id,
    el: plain,
    fullText: initialText || "",
    index: 0,
    timer: null,
    lastRender: 0,
    onDone: onDone || null,
    completeOnSkip,
    cancelled: false,
  };
  typewriterControllers.set(id, controller);
  queueTypewriterWork(controller);
  return controller;
}

function queueTypewriterWork(controller) {
  if (!controller) return;
  if (controller.timer) return;
  controller.timer = setInterval(() => stepTypewriter(controller), TYPEWRITER_SPEED_MS);
}

function stepTypewriter(controller) {
  if (!controller || controller.cancelled) {
    clearInterval(controller?.timer);
    return;
  }
  const target = controller.fullText || "";
  if (controller.index >= target.length) {
    clearInterval(controller.timer);
    controller.timer = null;
    if (controller.onDone) controller.onDone();
    return;
  }
  const now = Date.now();
  if (now - controller.lastRender < TYPEWRITER_DEBOUNCE_MS && target.length - controller.index > 50) {
    controller.index += Math.max(2, Math.floor((target.length - controller.index) / 40));
  } else {
    controller.index += Math.max(1, Math.floor(target.length > 400 ? 3 : 2));
    controller.lastRender = now;
  }
  const slice = target.slice(0, Math.min(controller.index, target.length));
  controller.el.textContent = slice;
}

function updateTypewriter(bubble, text) {
  const controller = getTypewriterController(bubble) || startTypewriter(bubble, text);
  controller.fullText = text;
  queueTypewriterWork(controller);
  const id = bubble?.dataset?.msgId;
  if (id) {
    const existingState = messageState.get(id) || { streamText: "", lastAnswer: "" };
    messageState.set(id, { ...existingState, streamText: text, lastAnswer: text });
  }
}

function finishTypewriter(bubble, { invokeDone = false } = {}) {
  const controller = getTypewriterController(bubble);
  if (!controller) return;
  controller.index = controller.fullText.length;
  controller.el.textContent = controller.fullText;
  if (controller.timer) clearInterval(controller.timer);
  controller.timer = null;
  if (invokeDone && controller.onDone) controller.onDone();
}

function cancelTypewriter(bubble) {
  const controller = getTypewriterController(bubble);
  if (!controller) return;
  controller.cancelled = true;
  if (controller.timer) clearInterval(controller.timer);
  typewriterControllers.delete(controller.id);
}

function skipTypewriter(bubble) {
  const controller = getTypewriterController(bubble);
  if (!controller) return;
  finishTypewriter(bubble, { invokeDone: controller.completeOnSkip });
}

function skipAllTypewriters() {
  typewriterControllers.forEach((controller) => {
    const bubble = [...chatLog.querySelectorAll(".bubble.assistant")].find(b => b.dataset.msgId === controller.id);
    if (bubble) skipTypewriter(bubble);
  });
}

function renderFinalAssistant(bubble, text, citationMap = new Map(), extras = {}) {
  if (!bubble) return;
  cancelTypewriter(bubble);
  const msgId = bubble.dataset.msgId;
  const state = msgId ? messageState.get(msgId) : null;
  const fallbackText = (text || "").trim() || (state?.lastAnswer || "").trim() || (state?.streamText || "").trim();
  const finalText = fallbackText || "(empty assistant message — formatting fallback engaged)";
  const pipeLines = finalText.split(/\n/).filter(line => /^\s*\|/.test(line)).length;
  const textBlock = ensureBubbleTextNode(bubble, { reset: true });
  const requestId = extras?.debug?.requestId || bubble?.dataset?.requestId || state?.requestId || null;
  const normalizedSources = normalizeCitedSources(Array.isArray(extras?.sources) ? extras.sources : [], citationMap);
  const preserveCitationMap = typeof extras?.preserveCitationMap === "boolean"
    ? extras.preserveCitationMap
    : citationMap instanceof Map && citationMap.size > 0;
  if (DEBUG_RENDER) {
    console.log("[RENDER_CHECK]", { msgId, answerLen: finalText.length, pipeLines });
    console.log("[RENDER_DEBUG]", {
      req: requestId,
      rawLen: finalText.length,
      table: looksLikeMarkdownTable(finalText),
      htmlTable: looksLikeHtmlTable(finalText),
    });
  }
  if (DEBUG_TABLE_RAW) {
    console.log("[RENDER_RAW_TEXT]", { msgId, requestId, text: finalText });
  }
  const renderPre = () => {
    const pre = document.createElement("pre");
    pre.className = "answer-pre";
    pre.textContent = finalText;
    const wrapper = decorateAssistantResponse("", citationMap, {
      msgId,
      answerText: finalText,
      sources: normalizedSources,
      preserveCitationMap,
    });
    const content = wrapper.querySelector(".response-content");
    if (content) content.appendChild(pre);
    return wrapper;
  };

  let node = null;
  try {
    node = decorateAssistantResponse(finalText, citationMap, {
      msgId,
      answerText: finalText,
      sources: normalizedSources,
      preserveCitationMap,
    });
    if (!node || ((node.innerText || "").trim().length === 0 && finalText.trim().length > 0)) {
      if (DEBUG_RENDER) console.log("[RENDER_DEBUG_EMPTY_BODY]", { req: requestId, msgId });
      node = renderPre();
    }
  } catch (err) {
    console.error("[RENDER_ERROR]", { msgId, err, answerLen: (finalText || "").length, requestId: extras?.debug?.requestId });
    node = renderPre();
  }
  textBlock.replaceChildren(node);
  const renderedNow = (textBlock.innerText || "").trim().length;
  if (renderedNow === 0 && finalText.trim().length > 0) {
    textBlock.replaceChildren(renderPre());
  }
  enhanceAutoTables(textBlock);
  enhanceTablesIn(textBlock);
  if (node && node.classList?.contains("response-card")) {
    finalizeResponseCard(node);
  }
  bubble.classList.toggle("has-response", Boolean(node && node.classList?.contains("response-card")));
  renderReferenceChips(bubble, extras.references);
  if (extras.evidence) {
    renderEvidenceToggle(bubble, extras.evidence);
  }
  renderMathIfReady(bubble);
  if (DEBUG_RENDER) {
    const renderedLen = (textBlock.innerText || "").length;
    const tableCount = textBlock.querySelectorAll("table").length;
    console.log("[RENDER_CHECK_DONE]", { msgId, answerLen: finalText.length, renderedLen, pipeLines, tableCount });
  }
  if (DEBUG_STREAMING) {
    console.log("[RENDER_FINAL]", {
      msgId,
      textLen: (text || "").length,
      finalLen: finalText.length,
      streamLen: state?.streamText?.length || 0,
      docId: bubble?.dataset?.docId || null,
      evidenceCount: extras.evidence?.length || 0,
    });
  }
  if (DEBUG_LECTURE_UI) {
    const container = bubble.querySelector(".bubble-text");
    const renderedLen = container?.innerText?.trim().length || 0;
    if (finalText.trim().length > 0 && renderedLen === 0) {
      const cs = container ? getComputedStyle(container) : null;
      console.error("[INV-3_FAIL]", {
        msgId,
        docId: bubble?.dataset?.docId || null,
        requestId: extras?.debug?.requestId,
        finalPreview: finalText.slice(0, 200),
        style: cs
          ? {
              display: cs.display,
              visibility: cs.visibility,
              opacity: cs.opacity,
              height: cs.height,
              overflow: cs.overflow,
            }
          : null,
      });
    }
  }
}

function renderWithTypewriter(bubble, text, citationMap = new Map(), extras = {}) {
  if (!bubble) return;
  const content = (text || "").trim();
  const msgId = bubble.dataset.msgId || generateMessageId();
  bubble.dataset.msgId = msgId;
  const existingState = messageState.get(msgId) || { streamText: "", lastAnswer: "" };
  messageState.set(msgId, { ...existingState, lastAnswer: content || existingState.lastAnswer, streamText: existingState.streamText || content });
  if (!content) {
    const fallback = appendChatMessage("assistant", "Failed! >:(\n\nLecture ask returned empty answer.", { track: false });
    fallback.classList.add("error");
    return;
  }
  startTypewriter(bubble, content, {
    completeOnSkip: true,
    onDone: () => renderFinalAssistant(bubble, content, citationMap, extras),
  });
}

function updateHashtagCloud(promptText) {
  const list = extractHashtags(promptText);
  if (hashtagCloud) {
    if (!list.length) {
      hashtagCloud.innerHTML = "";
    } else {
      hashtagCloud.innerHTML = list
        .map(tag => `<span class="hashtag-chip">${escapeHTML(tag)}</span>`)
        .join("");
    }
  }
  return list;
}

function normalizeTopicEntries(input) {
  if (!Array.isArray(input)) return [];
  return input
    .map((entry) => {
      if (typeof entry === "string") {
        const text = entry.trim();
        return text ? { text, count: 1, aliases: [] } : null;
      }
      if (entry && typeof entry === "object") {
        const text = (entry.text || entry.displayPhrase || entry.label || entry.tag || "").toString().trim();
        const count = Number.isFinite(entry.count) ? Math.max(1, entry.count) : 1;
        const aliases = Array.isArray(entry.aliases)
          ? entry.aliases.map(alias => (alias || "").toString().trim()).filter(Boolean).slice(0, 5)
          : [];
        if (!text) return null;
        return { text, count, aliases };
      }
      return null;
    })
    .filter(Boolean);
}

function renderMetaTopicsChart(entries, options = {}) {
  if (!metaTopicsChart) return;
  const list = normalizeTopicEntries(entries);
  if (!list.length) {
    metaTopicsChart.dataset.empty = "true";
    const emptyMessage = typeof options.emptyMessage === "string" && options.emptyMessage.trim()
      ? options.emptyMessage
      : "No topic data yet.";
    metaTopicsChart.textContent = emptyMessage;
    return;
  }
  const sorted = list
    .slice()
    .sort((a, b) => (b.count - a.count) || a.text.localeCompare(b.text));
  const maxCount = Math.max(1, ...sorted.map(entry => entry.count));
  metaTopicsChart.dataset.empty = "false";
  metaTopicsChart.innerHTML = sorted
    .slice(0, 12)
    .map(entry => {
      const pct = Math.max(1, Math.round((entry.count / maxCount) * 100));
      const label = escapeHTML(entry.text);
      const aliases = Array.isArray(entry.aliases) ? entry.aliases : [];
      const aliasText = aliases
        .filter(alias => alias && alias.toLowerCase() !== entry.text.toLowerCase())
        .join(" · ");
      const aliasMarkup = aliasText
        ? `<div class="meta-topic-aliases" title="${escapeHTML(aliasText)}">${escapeHTML(aliasText)}</div>`
        : "";
      return `
        <div class="meta-topic-row">
          <div class="meta-topic-label-wrap">
            <div class="meta-topic-label" title="${label}">${label}</div>
            ${aliasMarkup}
            <div class="meta-topic-track" aria-hidden="true">
              <div class="meta-topic-fill" style="width:${pct}%"></div>
            </div>
          </div>
          <div class="meta-topic-count">${entry.count}</div>
        </div>
      `;
    })
    .join("");
}

const META_BODY_TEMPLATE = `
  <div class="meta-group">
    <label class="meta-label" for="metaLectureSelect">Lecture</label>
    <select id="metaLectureSelect">
      <option value="">Select a lecture...</option>
    </select>
    <div class="meta-helper">Pick a lecture to view top topics.</div>
  </div>
  <div class="meta-group">
    <div class="meta-section-title">Top Topics</div>
    <div id="metaTopicsChart" class="meta-topics-chart" data-empty="true">Select a lecture to view top topics.</div>
  </div>
  <div id="metaError" class="meta-error" role="status" aria-live="polite" hidden></div>
`;

function mountMetaBody() {
  if (!metaBody || metaBody.childElementCount) return;
  metaBody.innerHTML = META_BODY_TEMPLATE;
  metaUpdated = metaUpdated || metaBody.closest(".analytics-card")?.querySelector("#metaUpdated") || null;
  metaLectureSelect = metaBody.querySelector("#metaLectureSelect");
  metaTopicsChart = metaBody.querySelector("#metaTopicsChart");
  metaError = metaBody.querySelector("#metaError");
  if (metaLectureSelect) {
    metaLectureSelect.addEventListener("change", () => refreshMetaDataPanel());
  }
}

function unmountMetaBody() {
  if (!metaBody) return;
  metaBody.replaceChildren();
  metaUpdated = null;
  metaLectureSelect = null;
  metaTopicsChart = null;
  metaError = null;
}

function applyMetaExpandedState(isExpanded) {
  metaDataExpanded = isExpanded;
  if (metaDataPanel) metaDataPanel.classList.toggle("is-collapsed", !isExpanded);
  if (metaToggle) {
    metaToggle.setAttribute("aria-label", isExpanded ? "Collapse meta data" : "Expand meta data");
  }
  if (metaBody) metaBody.hidden = !isExpanded;
}

function setMetaUnlockError(message) {
  if (!metaUnlockError) return;
  if (!message) {
    metaUnlockError.hidden = true;
    metaUnlockError.textContent = "";
    return;
  }
  metaUnlockError.hidden = false;
  metaUnlockError.textContent = message;
}

function formatMetaTimestamp(value) {
  if (!value) return "—";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "—";
  return parsed.toLocaleString();
}

function setMetaUpdated(value) {
  if (!metaUpdated) return;
  metaUpdated.textContent = formatMetaTimestamp(value);
}

function setMetaErrorMessage(message) {
  if (!metaError) return;
  if (!message) {
    metaError.hidden = true;
    metaError.textContent = "";
    return;
  }
  metaError.hidden = false;
  metaError.textContent = message;
}

function resetMetaPanel() {
  setMetaUpdated("");
  setMetaErrorMessage("");
  renderMetaTopicsChart([], { emptyMessage: "Select a lecture to view top topics." });
}

function applyMetaAnalyticsPayload(payload, docId, docTitle) {
  if (!payload || !docId) return false;
  const payloadDocId = typeof payload?.docId === "string"
    ? payload.docId
    : typeof payload?.conversation_id === "string"
      ? payload.conversation_id
      : "";
  const payloadTitle = typeof payload?.title === "string"
    ? payload.title
    : typeof payload?.docTitle === "string"
      ? payload.docTitle
      : "";
  if (payloadDocId) {
    if (payloadDocId !== docId) return false;
  } else if (docTitle) {
    if (!payloadTitle || payloadTitle !== docTitle) return false;
  } else {
    return false;
  }
  renderMetaTopicsChart(Array.isArray(payload?.topics) ? payload.topics : []);
  setMetaUpdated(payload.updated_at || "");
  setMetaErrorMessage("");
  return true;
}

function getSelectedMetaDoc() {
  if (!metaLectureSelect) return null;
  const docId = metaLectureSelect.value;
  if (!docId) return null;
  return libraryListItems.find(item => item.docId === docId) || {
    docId,
    title: metaLectureSelect.options[metaLectureSelect.selectedIndex]?.textContent || docId,
  };
}

function refreshMetaDataPanel() {
  if (!metaLectureSelect) return;
  if (!metaDataUnlocked) {
    resetMetaPanel();
    return;
  }
  const doc = getSelectedMetaDoc();
  if (!doc) {
    resetMetaPanel();
    return;
  }
  fetchMetaAnalytics(doc);
}

async function fetchMetaAnalytics(doc) {
  if (!metaDataUnlocked || !doc?.docId) return;
  const requestId = ++metaAnalyticsRequestId;
  if (metaAnalyticsController) metaAnalyticsController.abort();
  const controller = new AbortController();
  metaAnalyticsController = controller;
  setMetaErrorMessage("");
  setMetaUpdated("");
  renderMetaTopicsChart([], { emptyMessage: "Loading topics..." });
  try {
    const resp = await fetch(`/api/admin/analytics?docId=${encodeURIComponent(doc.docId)}`, { signal: controller.signal });
    const payload = await resp.json().catch(() => ({}));
    if (controller.signal.aborted || requestId !== metaAnalyticsRequestId) return;
    if (!resp.ok) {
      throw new Error(payload?.error || payload?.details || "Meta data fetch failed.");
    }
    if (!metaDataUnlocked) {
      return;
    }
    if (!applyMetaAnalyticsPayload(payload, doc.docId, getLectureDisplayLabel(doc) || doc.title || doc.key || doc.docId)) {
      throw new Error("Meta data response did not match the selected lecture.");
    }
  } catch (error) {
    if (controller.signal.aborted) return;
    console.warn("Meta data fetch failed", error);
    renderMetaTopicsChart([], { emptyMessage: "No topic data yet." });
    setMetaUpdated("");
    setMetaErrorMessage(error instanceof Error ? error.message : "Meta data fetch failed.");
  }
}

async function recordLectureAnalytics(doc) {
  if (!metaDataUnlocked || !doc?.docId) return;
  await fetchMetaAnalytics(doc);
}

function updateMetaLectureOptions(state = "ready") {
  if (!metaLectureSelect) return;
  const currentValue = metaLectureSelect.value;
  if (state === "loading") {
    metaLectureSelect.innerHTML = `<option value="">Loading lectures...</option>`;
    return;
  }
  if (state === "error") {
    metaLectureSelect.innerHTML = `<option value="">Failed to load</option>`;
    resetMetaPanel();
    return;
  }
  applyLectureDisplayLabels(libraryListItems);
  metaLectureSelect.innerHTML = `<option value="">Select a lecture...</option>` + libraryListItems.map(item => {
    const label = getLectureDisplayLabel(item) || item.title || item.key || item.docId;
    return `<option value="${escapeHTML(item.docId)}">${escapeHTML(label)}</option>`;
  }).join("");
  if (currentValue && libraryListItems.some(item => item.docId === currentValue)) {
    metaLectureSelect.value = currentValue;
  } else if (librarySelect?.value && libraryListItems.some(item => item.docId === librarySelect.value)) {
    metaLectureSelect.value = librarySelect.value;
  } else {
    metaLectureSelect.value = "";
  }
  refreshMetaDataPanel();
}

function updateMachineLectureOptions(state = "ready") {
  if (!machineLectureSelect) return;
  const currentValue = machineLectureSelect.value;
  if (state === "loading") {
    machineLectureSelect.innerHTML = `<option value="">Loading lectures...</option>`;
    return;
  }
  if (state === "error") {
    machineLectureSelect.innerHTML = `<option value="">Failed to load</option>`;
    setStatus(machineStatus, "error", "Failed to load lectures.");
    if (machineResult) machineResult.innerHTML = "";
    return;
  }
  applyLectureDisplayLabels(libraryListItems);
  machineLectureSelect.innerHTML = `<option value="">Select a lecture...</option>` + libraryListItems.map(item => {
    const label = getLectureDisplayLabel(item) || item.title || item.key || item.docId;
    return `<option value="${escapeHTML(item.docId)}">${escapeHTML(label)}</option>`;
  }).join("");
  if (currentValue && libraryListItems.some(item => item.docId === currentValue)) {
    machineLectureSelect.value = currentValue;
  } else {
    machineLectureSelect.value = "";
  }
  updateMachineSelection();
}

function updateMachineSelection() {
  if (!machineLectureSelect) return;
  const previousDocId = activeMachineDoc?.docId || "";
  const docId = machineLectureSelect.value;
  const selectionChanged = docId !== previousDocId;
  activeMachineDoc = docId ? (libraryListItems.find(item => item.docId === docId) || null) : null;
  if (selectionChanged && machineResult) {
    machineResult.innerHTML = "";
  }
  if (selectionChanged && machineTxtDetails) {
    machineTxtDetails.innerHTML = "";
  }
  if (machineGenerateBtn) {
    machineGenerateBtn.disabled = !activeMachineDoc || machineGenerating;
  }
  if (!activeMachineDoc) {
    setStatus(machineStatus, "info", "Select a lecture and generate TXT.");
    updateMachineStudyState();
    return;
  }
  setStatus(machineStatus, "info", "Select a lecture and generate TXT.");
  updateMachineStudyState();
}

async function runMachineTxtGeneration() {
  if (!activeMachineDoc || machineGenerating) return;
  machineGenerating = true;
  if (machineGenerateBtn) machineGenerateBtn.disabled = true;
  setStatus(machineStatus, "pending", "Generating…");
  if (machineResult) machineResult.innerHTML = "";
  if (machineTxtDetails) machineTxtDetails.innerHTML = "";
  try {
    const docId = activeMachineDoc.docId;
    const lectureTitle = activeMachineDoc.title || activeMachineDoc.key || docId;
    const resp = await fetch("/api/machine/lecture-to-txt", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ docId, lectureTitle }),
    });
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok || payload?.ok === false) {
      const msg = payload?.error || payload?.message || "Failed to generate TXT.";
      throw new Error(msg);
    }
    const filename = payload?.displayName || payload?.filename || getLectureTxtDisplayName(lectureTitle);
    const storedKey = payload?.storedKey || "";
    const downloadUrl = payload?.downloadUrl || (storedKey
      ? `/api/machine/download?key=${encodeURIComponent(storedKey)}&filename=${encodeURIComponent(filename)}`
      : "");
    lastMachineTxt = {
      filename,
      storedKey,
      docId,
      lectureTitle,
      text: typeof payload?.text === "string" && payload.text ? payload.text : "",
    };
    setStatus(machineStatus, "success", filename);
    setPipelineLastRun(4);
    if (machineResult) {
      const downloadLink = downloadUrl
        ? `<a class="machine-download-link" href="${downloadUrl}" download="${escapeHTML(filename)}">Download TXT</a>`
        : "";
      machineResult.innerHTML = `
        <div class="machine-output-row">
          <div class="machine-meta">
            ${downloadLink}
          </div>
        </div>
      `.trim();
    }
    if (machineTxtDetails) {
      machineTxtDetails.innerHTML = storedKey
        ? `<div class="machine-key">Stored key: <code>${escapeHTML(storedKey)}</code></div>`
        : "";
    }
    if (machineUseLastToggle) {
      machineUseLastToggle.disabled = false;
      if (!machineTxtInput?.files?.length) {
        machineUseLastToggle.checked = true;
      }
    }
    updateMachineStudyState();
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to generate TXT.";
    setStatus(machineStatus, "error", message);
  } finally {
    machineGenerating = false;
    if (machineGenerateBtn) machineGenerateBtn.disabled = !activeMachineDoc;
  }
}

function getMachineTxtSource() {
  const useLast = Boolean(machineUseLastToggle?.checked && lastMachineTxt);
  if (useLast && lastMachineTxt) {
    return {
      type: "last",
      filename: lastMachineTxt.filename,
      storedKey: lastMachineTxt.storedKey,
      docId: lastMachineTxt.docId,
      lectureTitle: lastMachineTxt.lectureTitle,
      text: lastMachineTxt.text,
    };
  }
  const file = machineTxtInput?.files?.[0] || null;
  if (file) {
    return { type: "file", filename: file.name, file };
  }
  return null;
}

function resolveStudyGuideDownloadUrl(storedKey, filename, fallbackUrl) {
  if (typeof fallbackUrl === "string" && fallbackUrl.trim()) return fallbackUrl;
  if (!storedKey) return "";
  const safeName = filename || "Study_Guide.html";
  return `/api/machine/download?key=${encodeURIComponent(storedKey)}&filename=${encodeURIComponent(safeName)}`;
}

function loadStoredStudyGuideOutput() {
  try {
    const raw = localStorage.getItem(STUDY_GUIDE_STATE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    const storedKey = typeof parsed.storedKey === "string" ? parsed.storedKey.trim() : "";
    if (!storedKey) return null;
    const filename = typeof parsed.filename === "string" && parsed.filename.trim()
      ? parsed.filename.trim()
      : "Study_Guide.html";
    const downloadUrl = resolveStudyGuideDownloadUrl(storedKey, filename, parsed.downloadUrl);
    return {
      ...parsed,
      storedKey,
      filename,
      downloadUrl,
    };
  } catch {
    return null;
  }
}

function persistStudyGuideOutput(output) {
  if (!output) {
    try {
      localStorage.removeItem(STUDY_GUIDE_STATE_KEY);
    } catch {}
    return;
  }
  const payload = {
    storedKey: output.storedKey || "",
    filename: output.filename || "",
    downloadUrl: output.downloadUrl || "",
    lectureTitle: output.lectureTitle || "",
    docId: output.docId || "",
    mode: output.mode || "",
    sourceStoredKey: output.sourceStoredKey || "",
  };
  try {
    localStorage.setItem(STUDY_GUIDE_STATE_KEY, JSON.stringify(payload));
  } catch {}
}

function loadStoredPublishResult() {
  try {
    const raw = localStorage.getItem(STUDY_GUIDE_PUBLISH_STATE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    if (typeof parsed.code !== "string" || !parsed.code.trim()) return null;
    return parsed;
  } catch {
    return null;
  }
}

function persistPublishResult(result) {
  if (!result) {
    try {
      localStorage.removeItem(STUDY_GUIDE_PUBLISH_STATE_KEY);
    } catch {}
    return;
  }
  try {
    localStorage.setItem(STUDY_GUIDE_PUBLISH_STATE_KEY, JSON.stringify(result));
  } catch {}
}

function hydrateStudyGuideState() {
  const storedOutput = loadStoredStudyGuideOutput();
  if (storedOutput) {
    lastStudyGuideOutput = storedOutput;
  }
  const storedPublish = loadStoredPublishResult();
  if (storedPublish && storedOutput?.storedKey) {
    const sourceKey = storedPublish.sourceKey || storedOutput.storedKey;
    if (sourceKey !== storedOutput.storedKey) {
      lastPublishResult = null;
      return;
    }
    lastPublishResult = { ...storedPublish, sourceKey };
    persistPublishResult(lastPublishResult);
    renderPublishResult(lastPublishResult);
    return;
  }
  lastPublishResult = null;
}

function updateMachineStudyState(opts = {}) {
  const preserveStatus = Boolean(opts.preserveStatus);
  const useLast = Boolean(machineUseLastToggle?.checked && lastMachineTxt);
  if (machineTxtInput) machineTxtInput.disabled = useLast;
  if (machineTxtButton) machineTxtButton.disabled = useLast;
  if (machineTxtName) {
    const file = machineTxtInput?.files?.[0] || null;
    const filename = useLast && lastMachineTxt
      ? lastMachineTxt.filename
      : file
        ? file.name
        : "No TXT selected";
    machineTxtName.textContent = filename;
    if (filename && filename !== "No TXT selected") {
      machineTxtName.title = filename;
    } else {
      machineTxtName.removeAttribute("title");
    }
  }
  const source = getMachineTxtSource();
  if (machineStudyBtn) {
    machineStudyBtn.disabled = !source || machineStudyGenerating;
  }
  if (!machineStudyStatus) return;
  if (!source) {
    setStatus(machineStudyStatus, "info", "Please upload or generate a TXT first.");
    if (machineStudyResult) machineStudyResult.innerHTML = "";
    return;
  }
  if (!preserveStatus) {
    setStatus(machineStudyStatus, "info", `Using TXT: ${source.filename}`);
  }
  updatePublishState();
}

async function loadLastMachineTxtText() {
  if (!lastMachineTxt) return "";
  if (lastMachineTxt.text) return lastMachineTxt.text;
  if (!lastMachineTxt.storedKey) return "";
  const url = `/api/machine/download?key=${encodeURIComponent(lastMachineTxt.storedKey)}&filename=${encodeURIComponent(lastMachineTxt.filename)}`;
  const resp = await fetch(url);
  if (!resp.ok) return "";
  const text = await resp.text();
  lastMachineTxt.text = text;
  return text;
}

function titleFromTxtFilename(name) {
  return cleanLectureLabel(name || "");
}

async function readJsonResponse(resp) {
  const text = await resp.text().catch(() => "");
  if (!text) return { payload: {}, rawText: "" };
  try {
    const parsed = JSON.parse(text);
    return { payload: parsed && typeof parsed === "object" ? parsed : {}, rawText: text };
  } catch {
    return { payload: {}, rawText: text };
  }
}

function extractErrorMessage(payload, rawText, fallback) {
  const message = payload?.error || payload?.message || "";
  const trimmed = (rawText || "").trim();
  let base = "";
  if (message && typeof message === "string") {
    base = message;
  } else if (trimmed) {
    base = trimmed.startsWith("<") ? fallback : trimmed.length > 200 ? `${trimmed.slice(0, 200)}...` : trimmed;
  } else {
    base = fallback;
  }
  const details = [];
  if (payload?.stage) details.push(`stage: ${payload.stage}`);
  if (payload?.debugCode) details.push(`code: ${payload.debugCode}`);
  const received = payload?.received;
  if (received && typeof received === "object") {
    const receivedParts = [];
    if (received.slidesPdf?.name) {
      receivedParts.push(`slidesPdf=${received.slidesPdf.name} (${received.slidesPdf.type || "unknown"}, ${received.slidesPdf.size || 0}B)`);
    }
    if (received.slideImagesFirst?.name) {
      receivedParts.push(`slideImages[0]=${received.slideImagesFirst.name} (${received.slideImagesFirst.type || "unknown"}, ${received.slideImagesFirst.size || 0}B)`);
    } else if (typeof received.slideImagesCount === "number") {
      receivedParts.push(`slideImages=${received.slideImagesCount}`);
    }
    if (received.transcriptTxt?.name) {
      receivedParts.push(`transcriptTxt=${received.transcriptTxt.name} (${received.transcriptTxt.type || "unknown"}, ${received.transcriptTxt.size || 0}B)`);
    }
    if (receivedParts.length) details.push(`received: ${receivedParts.join("; ")}`);
  }
  if (details.length) {
    return `${base} (${details.join(", ")})`;
  }
  return base;
}

// Manual test checklist:
// - Generate study guide succeeds and shows download link.
// - Publish returns a code and keeps publish status independent of generation.
// - Refresh keeps the last study guide output and publish code.
// - Student retrieve works with the published code.
async function runMachineStudyGuideGeneration() {
  if (machineStudyGenerating) return;
  const source = getMachineTxtSource();
  if (!source) {
    setStatus(machineStudyStatus, "error", "Please upload or generate a TXT first.");
    return;
  }
  const mode = machineStudyMode?.value === "canonical" ? "canonical" : "maximal";
  let publishLectureTitle = "";
  let publishDocId = "";
  const previousStoredKey = lastStudyGuideOutput?.storedKey || "";
  machineStudyGenerating = true;
  if (machineStudyBtn) machineStudyBtn.disabled = true;
  updatePublishState();
  setStatus(machineStudyStatus, "pending", "Generating…");
  if (machineStudyResult) machineStudyResult.innerHTML = "";
  try {
    let resp;
    if (source.type === "file") {
      const form = new FormData();
      form.append("txtFile", source.file, source.filename);
      const lectureTitle = titleFromTxtFilename(source.filename);
      if (lectureTitle) form.append("lectureTitle", lectureTitle);
      if (activeMachineDoc?.docId) form.append("docId", activeMachineDoc.docId);
      form.append("mode", mode);
      publishLectureTitle = lectureTitle;
      publishDocId = activeMachineDoc?.docId || "";
      resp = await fetch("/api/machine/generate-study-guide", { method: "POST", body: form });
    } else {
      const lectureTitle = source.lectureTitle || activeMachineDoc?.title || "";
      const docId = source.docId || activeMachineDoc?.docId || "";
      publishLectureTitle = lectureTitle;
      publishDocId = docId;
      if (source.type === "last" && source.storedKey) {
        resp = await fetch("/api/machine/generate-study-guide", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({
            storedKey: source.storedKey,
            filename: source.filename,
            lectureTitle,
            docId,
            mode,
          }),
        });
      } else {
        const text = await loadLastMachineTxtText();
        if (!text) {
          throw new Error("Please upload or generate a TXT first.");
        }
        resp = await fetch("/api/machine/generate-study-guide", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ txt: text, lectureTitle, docId, mode }),
        });
      }
    }
    const { payload, rawText } = await readJsonResponse(resp);
    if (!resp.ok || payload?.ok === false) {
      const msg = extractErrorMessage(payload, rawText, "Failed to generate study guide.");
      throw new Error(msg);
    }
    const filename = payload?.filename || "Study_Guide_Lecture.html";
    const storedKey = payload?.storedKey || "";
    if (!storedKey) {
      throw new Error("Study guide generation returned no storedKey.");
    }
    const downloadUrl = resolveStudyGuideDownloadUrl(storedKey, filename, payload?.downloadUrl);
    setStatus(machineStudyStatus, "success", "Study guide ready");
    lastStudyGuideOutput = {
      filename,
      storedKey,
      downloadUrl,
      lectureTitle: publishLectureTitle,
      docId: publishDocId,
      mode,
      sourceStoredKey: payload?.sourceStoredKey || "",
    };
    if (previousStoredKey && storedKey && previousStoredKey !== storedKey) {
      clearPublishResult();
    }
    persistStudyGuideOutput(lastStudyGuideOutput);
    updatePublishState();
    setPipelineLastRun(5);
    if (machineStudyResult) {
      const downloadLink = downloadUrl
        ? `<a class="machine-download-link" href="${downloadUrl}" download="${escapeHTML(filename)}">Download Study Guide</a>`
        : "";
      const modeLabel = mode === "canonical" ? "Canonical (strict)" : "Maximal (exam-forward)";
      const modelLine = "<div class=\"machine-key\">Model: <code>gpt-4o</code></div>";
      const modeLine = `<div class=\"machine-key\">Mode: <code>${escapeHTML(modeLabel)}</code></div>`;
      const keyBlock = storedKey
        ? `<div class="machine-key">Stored key: <code>${escapeHTML(storedKey)}</code></div>`
        : "";
      machineStudyResult.innerHTML = `
        <div class="machine-output-row">
          <div class="machine-meta">
            ${downloadLink}
          </div>
        </div>
        ${modelLine}
        ${modeLine}
        ${keyBlock}
      `.trim();
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to generate study guide.";
    setStatus(machineStudyStatus, "error", message);
  } finally {
    machineStudyGenerating = false;
    updateMachineStudyState({ preserveStatus: true });
  }
}

const LOCK_ICON_LOCKED = `
  <svg class="lock-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false" stroke="currentColor" fill="none">
    <rect x="3.5" y="11" width="17" height="10" rx="2" ry="2"></rect>
    <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
  </svg>
`;

const LOCK_ICON_UNLOCKED = `
  <svg class="lock-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false" stroke="currentColor" fill="none">
    <rect x="3.5" y="11" width="17" height="10" rx="2" ry="2"></rect>
    <path d="M7 11V7a5 5 0 0 1 9.5-2"></path>
  </svg>
`;

function setLockIcon(button, isUnlocked) {
  if (!button) return;
  button.dataset.state = isUnlocked ? "unlocked" : "locked";
  button.innerHTML = isUnlocked ? LOCK_ICON_UNLOCKED : LOCK_ICON_LOCKED;
}

function applyMetaUnlockState(isUnlocked) {
  metaDataUnlocked = isUnlocked;
  if (metaDataPanel) metaDataPanel.classList.toggle("is-locked", !isUnlocked);
  if (metaLockBtn) {
    setLockIcon(metaLockBtn, isUnlocked);
    metaLockBtn.setAttribute("aria-label", isUnlocked ? "Lock meta data" : "Unlock meta data");
  }
  if (metaToggle) metaToggle.hidden = !isUnlocked;
  if (metaUnlockInput) metaUnlockInput.hidden = isUnlocked;
  setMetaUnlockError("");
  try {
    localStorage.setItem(META_UNLOCK_KEY, String(isUnlocked));
  } catch {}
  if (!isUnlocked) {
    if (metaAnalyticsController) {
      metaAnalyticsController.abort();
      metaAnalyticsController = null;
    }
    metaAnalyticsRequestId += 1;
    resetMetaPanel();
    unmountMetaBody();
    applyMetaExpandedState(false);
  } else {
    mountMetaBody();
    resetMetaPanel();
    applyMetaExpandedState(true);
    updateMetaLectureOptions();
    refreshMetaDataPanel();
  }
}

function flashMachineUnlockError() {
  if (!machineUnlockInput) return;
  machineUnlockInput.classList.add("is-error");
  if (machineUnlockErrorTimer) {
    clearTimeout(machineUnlockErrorTimer);
  }
  machineUnlockErrorTimer = setTimeout(() => {
    machineUnlockInput.classList.remove("is-error");
    machineUnlockErrorTimer = null;
  }, 1000);
}

function applyMachineUnlockState(isUnlocked) {
  machineUnlocked = isUnlocked;
  if (machinePanel) machinePanel.classList.toggle("is-locked", !isUnlocked);
  if (machineLockBtn) {
    setLockIcon(machineLockBtn, isUnlocked);
    machineLockBtn.setAttribute("aria-label", isUnlocked ? "Lock The Machine" : "Unlock The Machine");
  }
  if (machineUnlockInput) {
    machineUnlockInput.classList.remove("is-error");
    machineUnlockInput.value = "";
  }
  try {
    localStorage.setItem(MACHINE_UNLOCK_KEY, String(isUnlocked));
  } catch {}
  if (!isUnlocked) {
    if (machinePanel) machinePanel.classList.add("is-collapsed");
    try {
      localStorage.setItem(MACHINE_COLLAPSE_KEY, "true");
    } catch {}
  }
}

function initMachineSubpanelToggle(panelId, toggleId, bodyId, storageKey) {
  const panel = document.getElementById(panelId);
  const toggle = document.getElementById(toggleId);
  const body = document.getElementById(bodyId);
  if (!panel || !toggle || !body) {
    console.warn("[SUBPANEL] missing nodes", { panelId, toggleId, bodyId });
    return false;
  }
  const initial = (() => {
    try {
      return localStorage.getItem(storageKey) === "true";
    } catch {
      return false;
    }
  })();
  const apply = (collapsed) => {
    panel.classList.toggle("is-collapsed", collapsed);
    body.hidden = collapsed;
    toggle.setAttribute("aria-expanded", String(!collapsed));
    try {
      localStorage.setItem(storageKey, String(collapsed));
    } catch {}
  };
  apply(initial);
  toggle.dataset.subpanelBound = "true";
  toggle.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const next = !panel.classList.contains("is-collapsed");
    apply(next);
  });
  return true;
}

const formatTime = () => new Intl.DateTimeFormat(undefined, {
  hour: "2-digit",
  minute: "2-digit",
}).format(new Date());

function formatRelativeTime(timestamp) {
  const ts = typeof timestamp === "number" ? timestamp : Date.now();
  const diff = Math.max(0, Date.now() - ts);
  const minute = 60 * 1000;
  const hour = 60 * minute;
  const day = 24 * hour;
  if (diff < minute) return "Just now";
  if (diff < hour) return `${Math.floor(diff / minute)}m ago`;
  if (diff < day) return `${Math.floor(diff / hour)}h ago`;
  if (diff < day * 2) return "Yesterday";
  if (diff < day * 7) return `${Math.floor(diff / day)}d ago`;
  return new Date(ts).toLocaleDateString();
}

function setPipelineLastRun(step) {
  const el = document.querySelector(`[data-role="last-run"][data-step="${step}"]`);
  if (!el) return;
  el.textContent = `Last run: ${new Date().toLocaleString()}`;
}

function getNumericStyleValue(raw) {
  const parsed = Number.parseFloat(raw);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getPipelineSectionNaturalHeight(section) {
  if (!section) return 0;
  const header = section.querySelector(".pipeline-card-header");
  const body = section.querySelector(".pipeline-subsection-body");
  const style = window.getComputedStyle(section);
  const gap = getNumericStyleValue(style.rowGap || style.gap);
  const paddingTop = getNumericStyleValue(style.paddingTop);
  const paddingBottom = getNumericStyleValue(style.paddingBottom);
  const borderTop = getNumericStyleValue(style.borderTopWidth);
  const borderBottom = getNumericStyleValue(style.borderBottomWidth);
  const headerHeight = header ? header.getBoundingClientRect().height : 0;
  const bodyHeight = body ? body.scrollHeight : 0;
  return headerHeight + bodyHeight + gap + paddingTop + paddingBottom + borderTop + borderBottom;
}

function syncPipelineDualCardHeight(root = document) {
  const card = root.querySelector(".pipeline-card--dual");
  if (!card) return;
  const sections = Array.from(card.querySelectorAll(".pipeline-subsection"));
  if (!sections.length) return;
  const heights = sections.map(getPipelineSectionNaturalHeight).filter(height => height > 0);
  if (!heights.length) return;
  const cardStyle = window.getComputedStyle(card);
  const paddingTop = getNumericStyleValue(cardStyle.paddingTop);
  const paddingBottom = getNumericStyleValue(cardStyle.paddingBottom);
  const borderTop = getNumericStyleValue(cardStyle.borderTopWidth);
  const borderBottom = getNumericStyleValue(cardStyle.borderBottomWidth);
  const target = Math.ceil(Math.max(...heights) + paddingTop + paddingBottom + borderTop + borderBottom);
  card.style.setProperty("--pipeline-dual-max-height", `${target}px`);
}

function updatePublishState() {
  if (!publishStudyGuideBtn || !publishStatus) return;
  const ready = Boolean(lastStudyGuideOutput?.storedKey);
  const isPublishing = publishInFlight;
  publishStudyGuideBtn.disabled = !ready || isPublishing;
  if (!ready) {
    setStatus(publishStatus, "info", "Generate a study guide to enable publishing.");
    return;
  }
  if (isPublishing) {
    setStatus(publishStatus, "pending", "Publishing...");
    return;
  }
  if (lastPublishResult && lastPublishResult.sourceKey === lastStudyGuideOutput?.storedKey) {
    setStatus(publishStatus, "success", "Published");
    return;
  }
  setStatus(publishStatus, "success", `Ready: ${lastStudyGuideOutput.filename}`);
}

function clearPublishResult() {
  lastPublishResult = null;
  persistPublishResult(null);
  if (publishResult) publishResult.innerHTML = "";
}

function renderPublishResult(result) {
  if (!publishResult) return;
  if (!result) {
    publishResult.innerHTML = "";
    return;
  }
  const code = result.code || "";
  const manifest = result.manifest || {};
  const publishedAt = manifest.publishedAt ? new Date(manifest.publishedAt).toLocaleString() : "";
  const detailLines = [];
  if (result.manifestKey) detailLines.push({ label: "Manifest", value: result.manifestKey });
  if (Array.isArray(result.publishedPaths)) {
    result.publishedPaths.forEach((path) => {
      if (path && typeof path === "string" && path !== result.manifestKey) {
        detailLines.push({ label: "Asset", value: path });
      }
    });
  }
  const detailsHtml = detailLines.length
    ? `<details class="publish-details">
        <summary>View details</summary>
        <div class="publish-details-body">
          ${detailLines.map(line => (
            `<div>${escapeHTML(line.label)}: <code>${escapeHTML(line.value)}</code></div>`
          )).join("")}
        </div>
      </details>`
    : "";
  publishResult.innerHTML = `
    <div class="publish-actions">
      <span class="status-tag ready">Published</span>
      ${publishedAt ? `<span>Published ${escapeHTML(publishedAt)}</span>` : ""}
    </div>
    <div class="publish-code-row">
      <span>Student Access Code</span>
      <span class="publish-code">${escapeHTML(code)}</span>
      <button type="button" class="btn btn-ghost" data-copy-code="${escapeHTML(code)}">Copy code</button>
    </div>
    ${detailsHtml}
  `.trim();
  const copyBtn = publishResult.querySelector("[data-copy-code]");
  if (copyBtn) {
    copyBtn.addEventListener("click", async () => {
      const text = copyBtn.getAttribute("data-copy-code") || "";
      const ok = await copyTextToClipboard(text);
      if (ok) {
        copyBtn.textContent = "Copied";
        setTimeout(() => { copyBtn.textContent = "Copy code"; }, 1200);
      }
    });
  }
}

async function runStudyGuidePublish() {
  if (publishInFlight) return;
  if (!lastStudyGuideOutput?.storedKey) {
    setStatus(publishStatus, "error", "Generate a study guide first.");
    return;
  }
  let failed = false;
  publishInFlight = true;
  updatePublishState();
  try {
    const payload = {
      storedKey: lastStudyGuideOutput.storedKey,
      filename: lastStudyGuideOutput.filename,
      lectureTitle: lastStudyGuideOutput.lectureTitle,
      docId: lastStudyGuideOutput.docId,
      mode: lastStudyGuideOutput.mode,
    };
    const request = {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify(payload),
    };
    let resp = await facultyFetch("/api/study-guides/publish", request, "study_guide_publish");
    if (resp.status === 404) {
      resp = await facultyFetch("/api/publish/study-guide", request, "study_guide_publish_fallback");
    }
    const { payload: data, rawText } = await readJsonResponse(resp);
    if (!resp.ok || data?.error) {
      const msg = extractErrorMessage(data, rawText, "Publish failed.");
      throw new Error(msg);
    }
    lastPublishResult = { ...data, sourceKey: lastStudyGuideOutput?.storedKey || "" };
    persistPublishResult(lastPublishResult);
    renderPublishResult(lastPublishResult);
    setStatus(publishStatus, "success", "Published");
  } catch (error) {
    failed = true;
    const message = error instanceof Error ? error.message : "Publish failed.";
    setStatus(publishStatus, "error", message);
  } finally {
    publishInFlight = false;
    if (failed) {
      if (publishStudyGuideBtn) publishStudyGuideBtn.disabled = false;
    } else {
      updatePublishState();
    }
  }
}

const ANKI_BADGE_LABELS = {
  EMPTY: "EMPTY",
  PDF_LOADED: "PDF LOADED",
  PARSING: "PARSING",
  IMAGES_READY: "IMAGES READY",
  TRANSCRIPT_READY: "TRANSCRIPT READY",
  READY_TO_GENERATE: "READY",
  GENERATING: "RUNNING",
  COMPLETE: "DONE",
  ERROR: "ERROR",
  READY: "READY",
  RUNNING: "RUNNING",
  DONE: "DONE",
};

function setAnkiBadgeState(state) {
  if (!ankiStatusBadge) return;
  const normalized = String(state || "READY").toUpperCase();
  const label = ANKI_BADGE_LABELS[normalized] || "READY";
  ankiStatusBadge.textContent = label;
  ankiStatusBadge.dataset.state = normalized.toLowerCase();
}

function setAnkiPdfStatus(state, message) {
  if (!ankiPdfStatus) return;
  if (!message) {
    ankiPdfStatus.hidden = true;
    ankiPdfStatus.textContent = "";
    ankiPdfStatus.removeAttribute("data-state");
    return;
  }
  ankiPdfStatus.hidden = false;
  setStatus(ankiPdfStatus, state, message);
}

function clearAnkiResult() {
  if (ankiGenerateResult) ankiGenerateResult.innerHTML = "";
  ankiCardsDownloadUrl = "";
  updateAnkiDownloads();
}

function clearAnkiSlidesZip() {
  if (ankiSlidesZipUrl) {
    URL.revokeObjectURL(ankiSlidesZipUrl);
  }
  ankiSlidesZipUrl = "";
  ankiSlidesZipName = "slides.zip";
}

function clearAnkiDerivedArtifacts() {
  clearAnkiSlidesZip();
  ankiSlidesOcrText = "";
  ankiManifest = null;
}

function updateAnkiDownloads() {
  if (ankiSlidesDownload) {
    if (ankiSlidesZipUrl) {
      ankiSlidesDownload.hidden = false;
      ankiSlidesDownload.href = ankiSlidesZipUrl;
      ankiSlidesDownload.setAttribute("download", ankiSlidesZipName || "slides.zip");
    } else {
      ankiSlidesDownload.hidden = true;
      ankiSlidesDownload.removeAttribute("href");
    }
  }
  if (ankiCardsDownload) {
    if (ankiCardsDownloadUrl) {
      ankiCardsDownload.hidden = false;
      ankiCardsDownload.href = ankiCardsDownloadUrl;
      ankiCardsDownload.setAttribute("download", "cards.tsv");
    } else {
      ankiCardsDownload.hidden = true;
      ankiCardsDownload.removeAttribute("href");
    }
  }
}

function buildAnkiLoadedRow(label, files, clearKey) {
  const pills = files.map((file) => {
    const name = typeof file === "string" ? file : (file?.name || "file");
    const tooltip = typeof file === "string" ? name : formatAnkiFileTooltip(file);
    return `<span class="machine-txt-pill" title="${escapeHTML(tooltip)}">${escapeHTML(name)}</span>`;
  }).join("");
  const clearButton = `
    <button type="button" class="btn btn-ghost btn-mini" data-anki-clear="${escapeHTML(clearKey)}">Clear</button>
  `;
  return `
    <div class="anki-loaded-row">
      <span class="anki-status-pill">${escapeHTML(label)}</span>
      <div class="anki-file-pills">${pills}</div>
      ${clearButton}
    </div>
  `.trim();
}

function renderAnkiLoadedList() {
  if (!ankiLoadedList) return;
  const rows = [];
  if (ankiPdfFile) {
    rows.push(buildAnkiLoadedRow("PDF loaded", [ankiPdfFile], "pdf"));
  }
  if (ankiImageFiles.length) {
    const label = ankiImagesSource === "pdf"
      ? "Slides source: JPEGs (converted from PDF)"
      : `${ankiImageFiles.length} images loaded`;
    rows.push(buildAnkiLoadedRow(label, ankiImageFiles, "images"));
  }
  if (ankiTranscriptFile) {
    rows.push(buildAnkiLoadedRow("Transcript loaded", [ankiTranscriptFile], "transcript"));
  }
  if (ankiBoldmapFile) {
    rows.push(buildAnkiLoadedRow("boldmap.csv loaded", [ankiBoldmapFile], "boldmap"));
  }
  if (ankiClassmateFile) {
    rows.push(buildAnkiLoadedRow("classmate_deck.tsv loaded", [ankiClassmateFile], "classmate"));
  }
  ankiLoadedList.innerHTML = rows.join("");
}

function setAnkiJobError(message, recovery) {
  ankiJobError = message || "Unexpected error.";
  ankiJobRecovery = recovery || "";
  ankiJobState = ANKI_JOB_STATES.ERROR;
  renderAnkiJobState({ preserveStatus: true });
}

function clearAnkiJobError() {
  ankiJobError = "";
  ankiJobRecovery = "";
  if (ankiJobState === ANKI_JOB_STATES.ERROR) {
    ankiJobState = ankiLastStableState;
  }
}

function deriveAnkiJobState() {
  if (ankiGenerating) return ANKI_JOB_STATES.GENERATING;
  if (ankiPdfConverting) return ANKI_JOB_STATES.PARSING;
  if (ankiCardsDownloadUrl) return ANKI_JOB_STATES.COMPLETE;
  const hasImages = ankiImageFiles.length > 0;
  const hasTranscript = Boolean(ankiTranscriptFile);
  if (hasImages && hasTranscript) return ANKI_JOB_STATES.READY_TO_GENERATE;
  if (hasTranscript) return ANKI_JOB_STATES.TRANSCRIPT_READY;
  if (hasImages) return ANKI_JOB_STATES.IMAGES_READY;
  if (ankiPdfFile) return ANKI_JOB_STATES.PDF_LOADED;
  return ANKI_JOB_STATES.EMPTY;
}

function renderAnkiJobState(opts = {}) {
  const preserveStatus = Boolean(opts.preserveStatus) || ankiGenerating || ankiPdfConverting;
  const effectiveState = ankiJobState === ANKI_JOB_STATES.ERROR ? ankiLastStableState : ankiJobState;
  const hasImages = ankiImageFiles.length > 0;
  const hasTranscript = Boolean(ankiTranscriptFile);
  const isLocked = ankiGenerating || ankiPdfConverting || ankiJobState === ANKI_JOB_STATES.COMPLETE;

  if (ankiPdfButton) {
    ankiPdfButton.disabled = effectiveState !== ANKI_JOB_STATES.EMPTY || isLocked;
  }
  if (ankiImagesButton) {
    ankiImagesButton.disabled = effectiveState !== ANKI_JOB_STATES.EMPTY || isLocked;
  }
  if (ankiParseButton) {
    ankiParseButton.disabled = effectiveState !== ANKI_JOB_STATES.PDF_LOADED || isLocked;
  }
  if (ankiTranscriptButton) {
    ankiTranscriptButton.disabled = !hasImages || isLocked;
  }
  if (ankiBoldmapButton) {
    ankiBoldmapButton.disabled = !hasImages || isLocked;
  }
  if (ankiClassmateButton) {
    ankiClassmateButton.disabled = !hasImages || isLocked;
  }
  if (ankiGenerateBtn) {
    ankiGenerateBtn.disabled = effectiveState !== ANKI_JOB_STATES.READY_TO_GENERATE || ankiGenerating;
  }

  setAnkiBadgeState(ankiJobState);
  updateAnkiDownloads();
  renderAnkiLoadedList();

  if (ankiGenerateStatus) {
    if (ankiJobState === ANKI_JOB_STATES.ERROR) {
      const recovery = ankiJobRecovery ? ` ${ankiJobRecovery}` : "";
      setStatus(ankiGenerateStatus, "error", `${ankiJobError}${recovery}`);
    } else if (!preserveStatus) {
      switch (ankiJobState) {
        case ANKI_JOB_STATES.EMPTY:
          setStatus(ankiGenerateStatus, "info", "Load a PDF or JPEGs to begin.");
          break;
        case ANKI_JOB_STATES.PDF_LOADED:
          setStatus(ankiGenerateStatus, "info", "PDF loaded. Click Parse to convert slides.");
          break;
        case ANKI_JOB_STATES.PARSING:
          setStatus(ankiGenerateStatus, "pending", "Parsing slides...");
          break;
        case ANKI_JOB_STATES.IMAGES_READY:
          setStatus(ankiGenerateStatus, "info", "Upload a transcript (.txt) to continue.");
          break;
        case ANKI_JOB_STATES.TRANSCRIPT_READY:
          setStatus(ankiGenerateStatus, "info", "Upload slide images or parse a PDF to continue.");
          break;
        case ANKI_JOB_STATES.READY_TO_GENERATE:
          setStatus(ankiGenerateStatus, "info", "Ready to generate.");
          break;
        case ANKI_JOB_STATES.COMPLETE:
          setStatus(ankiGenerateStatus, "success", "Anki deck ready.");
          break;
        default:
          break;
      }
    }
  }

  if (!ankiPdfConverting && ankiPdfStatus && ankiJobState !== ANKI_JOB_STATES.ERROR) {
    if (effectiveState === ANKI_JOB_STATES.PDF_LOADED) {
      setAnkiPdfStatus("info", "PDF loaded. Click Parse to convert slides.");
    } else if (!ankiPdfFile && !ankiImageFiles.length) {
      setAnkiPdfStatus("", "");
    }
  }
}

function updateAnkiGenerateState(opts = {}) {
  if (ankiJobState !== ANKI_JOB_STATES.ERROR) {
    const next = deriveAnkiJobState();
    if (next !== ankiJobState) {
      ankiJobState = next;
      if (next !== ANKI_JOB_STATES.PARSING && next !== ANKI_JOB_STATES.GENERATING) {
        ankiLastStableState = next;
      }
    }
  }
  renderAnkiJobState(opts);
}

function resolveAnkiLectureIdForManifest() {
  const directId = activeLibraryDoc?.docId || activeConversationMeta?.selectedDocId || "";
  const titleFallback = cleanLectureLabel(ankiPdfFile?.name || ankiTranscriptFile?.name || "");
  return directId || titleFallback || "lecture";
}

async function buildAnkiSlidesZip(files) {
  const JSZip = await loadJsZip();
  const zip = new JSZip();
  files.forEach((file) => {
    zip.file(file.name, file);
  });
  return zip.generateAsync({ type: "blob", compression: "DEFLATE" });
}

async function buildAnkiManifest(images, opts = {}) {
  const indexed = images.map((file, index) => ({ file, index }));
  const entries = new Array(images.length);
  await promisePool(indexed, 4, async (entry) => {
    const buffer = await entry.file.arrayBuffer();
    const hash = await computeSha256(buffer);
    entries[entry.index] = {
      page: entry.index + 1,
      basename: entry.file.name,
      bytes: entry.file.size || 0,
      sha256: hash,
    };
  });
  return {
    lecture_id: opts.lectureId || resolveAnkiLectureIdForManifest(),
    pdf_filename: opts.pdfFilename || null,
    page_count: images.length,
    images: entries,
  };
}

async function runAnkiOcr(images, opts = {}) {
  if (!images.length) return "";
  const total = images.length;
  const fileHash = opts.fileHash || createAnkiRequestId();
  const indexed = images.map((file, index) => ({ file, index }));
  const pages = new Array(total);
  await promisePool(indexed, OCR_CONCURRENCY, async (entry) => {
    if (opts.showProgress !== false) {
      setAnkiPdfStatus("pending", `OCR page ${entry.index + 1} / ${total}`);
    }
    try {
      const result = await ocrPageImage({
        fileHash,
        fileId: "",
        pageIndex: entry.index,
        blob: entry.file,
        filename: entry.file.name,
        signal: opts.signal,
      });
      pages[entry.index] = result?.text || "";
    } catch (err) {
      const message = err instanceof Error ? err.message : "OCR failed.";
      pages[entry.index] = `OCR failed: ${message}`;
    }
  });
  return pages.map((text, index) => (
    `=== Slide ${index + 1} | Image:${images[index].name} ===\n${String(text || "").trim()}`
  )).join("\n\n");
}

async function buildAnkiDerivedArtifacts(opts = {}) {
  const images = opts.images || [];
  if (!images.length) return;
  clearAnkiSlidesZip();
  const zipBlob = await buildAnkiSlidesZip(images);
  ankiSlidesZipUrl = URL.createObjectURL(zipBlob);
  ankiSlidesZipName = "slides.zip";
  ankiManifest = await buildAnkiManifest(images, {
    lectureId: opts.lectureId,
    pdfFilename: opts.pdfFilename,
  });
  if (opts.ocrMode === "required") {
    ankiSlidesOcrText = await runAnkiOcr(images, {
      fileHash: opts.fileHash,
      signal: opts.signal,
      showProgress: true,
    });
  } else if (opts.ocrMode === "background") {
    ankiSlidesOcrText = "OCR pending.";
    runAnkiOcr(images, {
      fileHash: opts.fileHash,
      signal: opts.signal,
      showProgress: false,
    })
      .then((text) => { ankiSlidesOcrText = text; })
      .catch((err) => {
        const message = err instanceof Error ? err.message : "OCR failed.";
        ankiSlidesOcrText = `OCR failed: ${message}`;
      });
  } else {
    ankiSlidesOcrText = "OCR not run.";
  }
  updateAnkiDownloads();
}

function handleAnkiClear(key) {
  let clearedSlides = false;
  if (key === "pdf") {
    ankiPdfFile = null;
    if (ankiImagesSource === "pdf") {
      ankiImageFiles = [];
      ankiImagesSource = "";
      clearedSlides = true;
    }
    ankiPdfConverting = false;
    setAnkiPdfStatus("", "");
    if (ankiPdfInput) ankiPdfInput.value = "";
  } else if (key === "images") {
    ankiImageFiles = [];
    ankiImagesSource = "";
    if (ankiImagesInput) ankiImagesInput.value = "";
    clearedSlides = true;
  } else if (key === "transcript") {
    ankiTranscriptFile = null;
    if (ankiTranscriptInput) ankiTranscriptInput.value = "";
  } else if (key === "boldmap") {
    ankiBoldmapFile = null;
    if (ankiBoldmapInput) ankiBoldmapInput.value = "";
  } else if (key === "classmate") {
    ankiClassmateFile = null;
    if (ankiClassmateInput) ankiClassmateInput.value = "";
  }
  if (clearedSlides) {
    clearAnkiDerivedArtifacts();
    setAnkiPdfStatus("", "");
  }
  clearAnkiResult();
  clearAnkiGenerateError();
  updateAnkiGenerateState();
}

async function runAnkiParse() {
  if (ankiPdfConverting) return;
  const file = ankiPdfFile;
  if (!file) {
    setAnkiJobError("Load a PDF before parsing.", "Use Load PDF, then click Parse.");
    return;
  }
  clearAnkiGenerateError();
  clearAnkiResult();
  clearAnkiDerivedArtifacts();
  ankiPdfConverting = true;
  updateAnkiGenerateState({ preserveStatus: true });
  if (ankiGenerateStatus) {
    setStatus(ankiGenerateStatus, "pending", "Parsing slides...");
  }

  let arrayBuffer = null;
  let fileHash = "";
  try {
    arrayBuffer = await file.arrayBuffer();
    fileHash = await computeSha256(arrayBuffer);
    const images = await convertPdfToAnkiImages(file, arrayBuffer);
    ankiImageFiles = images;
    ankiImagesSource = "pdf";
    updateAnkiGenerateState({ preserveStatus: true });
    try {
      await buildAnkiDerivedArtifacts({
        images,
        pdfFilename: file.name,
        fileHash,
        ocrMode: "required",
      });
      setAnkiPdfStatus("success", `Converted ${images.length} pages.`);
    } catch (err) {
      const message = err instanceof Error ? err.message : "Failed to finalize parse artifacts.";
      setAnkiJobError(message, "Try parsing again or re-upload the PDF.");
      setAnkiPdfStatus("error", message);
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : "Failed to parse PDF.";
    ankiImageFiles = [];
    ankiImagesSource = "";
    setAnkiJobError(message, "Try a different PDF or re-upload.");
    setAnkiPdfStatus("error", message);
  } finally {
    ankiPdfConverting = false;
    updateAnkiGenerateState();
  }
}

async function convertPdfToAnkiImages(file, arrayBufferOverride) {
  const arrayBuffer = arrayBufferOverride || await file.arrayBuffer();
  const pdfjsLib = await loadPdfJsLib();
  const pdfDoc = await pdfjsLib.getDocument({
    data: arrayBuffer,
    disableWorker: true,
  }).promise;
  const totalPages = pdfDoc?.numPages || 0;
  if (!totalPages) {
    throw new Error("PDF has no pages.");
  }
  if (totalPages > ANKI_MAX_PAGES) {
    throw new Error(`PDF has ${totalPages} pages. Max is ${ANKI_MAX_PAGES}.`);
  }
  const images = [];
  try {
    for (let pageNumber = 1; pageNumber <= totalPages; pageNumber += 1) {
      setAnkiPdfStatus("pending", `Converting PDF → images (page ${pageNumber} / ${totalPages})`);
      const render = await renderPdfPageToBlob(pdfjsLib, pdfDoc, pageNumber);
      const filename = `${pageNumber}.jpg`;
      images.push(new File([render.blob], filename, { type: "image/jpeg" }));
    }
  } finally {
    pdfDoc.cleanup?.();
    pdfDoc.destroy?.();
  }
  return images;
}

function renderAnkiResult(payload) {
  if (!ankiGenerateResult) return;
  if (!payload) {
    ankiGenerateResult.innerHTML = "";
    return;
  }
  const code = payload.code || "";
  const downloadUrl = payload.downloadUrl || "";
  ankiCardsDownloadUrl = downloadUrl;
  updateAnkiDownloads();
  const downloadLink = downloadUrl
    ? `<a class="machine-download-link" href="${escapeHTML(downloadUrl)}" download="cards.tsv">Download cards.tsv</a>`
    : "";
  ankiGenerateResult.innerHTML = `
    <div class="machine-output-row">
      <div class="machine-meta">
        ${downloadLink}
      </div>
    </div>
    <div class="anki-code-row">
      <span>Code:</span>
      <code>${escapeHTML(code)}</code>
      <button
        type="button"
        class="library-code-copy"
        data-anki-copy="${escapeHTML(code)}"
        aria-label="Copy retrieval code"
        title="Copy retrieval code"
      >
        <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <rect x="9" y="9" width="11" height="11" rx="2"></rect>
          <rect x="4" y="4" width="11" height="11" rx="2"></rect>
        </svg>
      </button>
      <span class="library-code-feedback" role="status" aria-live="polite"></span>
    </div>
    <div class="anki-note">Saved to owen-ingest/Anki Decks/</div>
  `.trim();
  const copyBtn = ankiGenerateResult.querySelector("[data-anki-copy]");
  const feedback = ankiGenerateResult.querySelector(".library-code-feedback");
  if (copyBtn) {
    copyBtn.addEventListener("click", async () => {
      const text = copyBtn.getAttribute("data-anki-copy") || "";
      const ok = await copyTextToClipboard(text);
      if (ok) {
        copyBtn.setAttribute("data-copied", "true");
        if (feedback) {
          feedback.textContent = "Copied";
          feedback.classList.add("is-visible");
        }
        setTimeout(() => {
          copyBtn.setAttribute("data-copied", "false");
          if (feedback) {
            feedback.textContent = "";
            feedback.classList.remove("is-visible");
          }
        }, 1200);
      }
    });
  }
}

async function runAnkiGeneration() {
  if (ankiGenerateAbortController) {
    ankiGenerateAbortController.abort();
  }
  const effectiveState = ankiJobState === ANKI_JOB_STATES.ERROR ? ankiLastStableState : ankiJobState;
  if (effectiveState !== ANKI_JOB_STATES.READY_TO_GENERATE) {
    setAnkiJobError("Not ready to generate.", "Upload slides and a transcript first.");
    return;
  }
  if (!ankiTranscriptFile || ankiImageFiles.length === 0) {
    setAnkiJobError("Upload transcript + slides before generating.", "Make sure images and transcript are loaded.");
    return;
  }
  clearAnkiGenerateError();
  clearAnkiResult();
  const requestId = createAnkiRequestId();
  const requestSeq = ankiGenerateRequestSeq + 1;
  ankiGenerateRequestSeq = requestSeq;
  ankiGenerateRequestId = requestId;
  const controller = new AbortController();
  ankiGenerateAbortController = controller;

  ankiGenerating = true;
  updateAnkiGenerateState({ preserveStatus: true });
  if (ankiGenerateStatus) {
    ankiGenerateStatus.dataset.state = "pending";
    ankiGenerateStatus.innerHTML = `Generating… <span class="anki-spinner" aria-hidden="true"></span>`;
  }

  const form = new FormData();
  const hasSlideImages = ankiImageFiles.length > 0;
  if (ankiPdfFile && !hasSlideImages) {
    form.append("slidesPdf", ankiPdfFile, ankiPdfFile.name);
  }
  if (hasSlideImages) {
    ankiImageFiles.forEach(file => form.append("slideImages[]", file, file.name));
  }
  form.append("transcriptTxt", ankiTranscriptFile, ankiTranscriptFile.name);
  if (ankiBoldmapFile) form.append("boldmapCsv", ankiBoldmapFile, ankiBoldmapFile.name);
  if (ankiClassmateFile) form.append("classmateDeckTsv", ankiClassmateFile, ankiClassmateFile.name);

  const titleFallback = cleanLectureLabel(ankiPdfFile?.name || ankiTranscriptFile?.name || "");
  const lectureTitle = (nameInput?.value || "").trim() || titleFallback;
  if (lectureTitle) form.append("lectureTitle", lectureTitle);
  const lectureId = activeLibraryDoc?.docId || activeConversationMeta?.selectedDocId || "";
  if (lectureId) form.append("lectureId", lectureId);

  try {
    const resp = await facultyFetch("/api/anki/generate", {
      method: "POST",
      body: form,
      signal: controller.signal,
      headers: { "x-request-id": requestId },
    }, "anki_generate");
    const { payload, rawText } = await readJsonResponse(resp);
    const responseRequestId = resolveAnkiResponseRequestId(payload, resp, requestId);
    if (requestSeq !== ankiGenerateRequestSeq) return;
    if (!resp.ok || payload?.error) {
      const msg = extractErrorMessage(payload, rawText, "Failed to generate Anki deck.");
      throw new Error(appendAnkiRequestId(msg, responseRequestId));
    }
    const code = payload?.code || "";
    const downloadUrl = payload?.downloadUrl || (code ? `/api/anki/download?code=${encodeURIComponent(code)}` : "");
    renderAnkiResult({ code, downloadUrl });
    setStatus(ankiGenerateStatus, "success", "Anki deck ready.");
    setPipelineLastRun(6);
  } catch (error) {
    if (requestSeq !== ankiGenerateRequestSeq) return;
    if (error?.name === "AbortError") return;
    const message = error instanceof Error ? error.message : "Failed to generate Anki deck.";
    setAnkiJobError(appendAnkiRequestId(message, ankiGenerateRequestId), "Fix the issue and try again.");
  } finally {
    if (requestSeq !== ankiGenerateRequestSeq) return;
    ankiGenerating = false;
    ankiGenerateAbortController = null;
    updateAnkiGenerateState();
  }
}

function setStatus(el, state, message) {
  if (!el) return;
  if (!message) {
    el.removeAttribute("data-state");
    el.textContent = "";
    return;
  }
  el.dataset.state = state || "info";
  el.textContent = message;
}

async function copyTextToClipboard(text) {
  if (!text) return false;
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    const fallback = document.createElement("textarea");
    fallback.value = text;
    fallback.style.position = "fixed";
    fallback.style.opacity = "0";
    document.body.appendChild(fallback);
    fallback.focus();
    fallback.select();
    try {
      return document.execCommand("copy");
    } catch {
      return false;
    } finally {
      fallback.remove();
    }
  }
}

function setUploadStatusState(state, message) {
  if (!uploadStatus) return;
  uploadStatus.classList.remove("upload-status");
  uploadStatus.removeAttribute("data-mode");
  uploadStatus.innerHTML = "";
  setStatus(uploadStatus, state, message);
}

function renderUploadSuccess(payload, displayName, statusState) {
  if (!uploadStatus) return;
  uploadStatus.classList.add("upload-status");
  uploadStatus.dataset.state = statusState || "success";
  uploadStatus.dataset.mode = "success";

  const nameSource = displayName || payload?.filename || getKeyLeaf(payload?.key || "");
  const cleanedFilename = cleanLectureLabel(nameSource) || nameSource || "Upload";
  const safeFilename = escapeHTML(cleanedFilename);
  const fullPath = payload?.bucket && payload?.key ? `${payload.bucket}/${payload.key}` : payload?.key || "";

  const vectorStatusRaw = payload?.vector_store_status || (payload?.vector_store_id ? "processing" : "");
  const vectorSummary = payload?.vector_store_id
    ? vectorStatusRaw === "completed"
      ? "ready"
      : vectorStatusRaw || "processing"
    : "unavailable";

  const ocrStatusRaw = payload?.ocr_status || (payload?.ocr_text_key ? "ready" : "");
  const ocrSummary = ocrStatusRaw
    ? ocrStatusRaw === "ready"
      ? "ready"
      : ocrStatusRaw === "error"
        ? "failed"
        : ocrStatusRaw === "empty"
          ? "no text extracted"
          : ocrStatusRaw
    : "";

  const warnings = [];
  if (payload?.vector_warning) warnings.push(payload.vector_warning);
  if (payload?.ocr_warning) warnings.push(payload.ocr_warning);

  const detailRows = [];
  if (fullPath) detailRows.push({ label: "Storage path", value: fullPath, copy: true, mono: true });
  if (payload?.ocr_text_key) detailRows.push({ label: "OCR text key", value: payload.ocr_text_key, mono: true });
  if (vectorSummary) detailRows.push({ label: "Vector index", value: vectorSummary });
  if (ocrSummary) detailRows.push({ label: "OCR", value: ocrSummary });
  if (warnings.length) detailRows.push({ label: "Warnings", value: warnings.join(" | ") });

  const detailsHtml = detailRows.length
    ? `<details class="upload-success-details">
        <summary>Details</summary>
        <div class="upload-success-details-body">
          ${detailRows.map((row) => {
            const valueHtml = row.mono
              ? `<code class="upload-success-path">${escapeHTML(row.value)}</code>`
              : `<span class="upload-success-detail-value">${escapeHTML(row.value)}</span>`;
            const copyBtn = row.copy
              ? `<button type="button" class="btn btn-ghost upload-success-copy" data-copy-path="${escapeHTML(row.value)}">Copy path</button>`
              : "";
            return `
              <div class="upload-success-detail-row">
                <span class="upload-success-detail-label">${escapeHTML(row.label)}</span>
                <span class="upload-success-detail-body">${valueHtml}</span>
                ${copyBtn}
              </div>`;
          }).join("")}
        </div>
      </details>`
    : "";

  uploadStatus.innerHTML = `
    <div class="upload-success-banner" role="status">
      <span class="upload-success-dot" aria-hidden="true"></span>
      <div class="upload-success-text">
        <div class="upload-success-primary">Upload complete</div>
        <div class="upload-success-secondary">Ready for indexing and OCR.</div>
      </div>
    </div>
    <div class="upload-success-filename" title="${safeFilename}">${safeFilename}</div>
    ${detailsHtml}
  `;

  const copyBtn = uploadStatus.querySelector("[data-copy-path]");
  if (copyBtn) {
    copyBtn.addEventListener("click", async () => {
      const path = copyBtn.getAttribute("data-copy-path") || "";
      if (!path) return;
      const ok = await copyTextToClipboard(path);
      if (!ok) return;
      const original = copyBtn.textContent || "Copy path";
      copyBtn.textContent = "Copied";
      setTimeout(() => {
        copyBtn.textContent = original;
      }, 1200);
    });
  }
}

function setLibraryAdvancedState(isOpen) {
  libraryAdvancedOpen = isOpen;
  if (libraryAdvancedBody) libraryAdvancedBody.classList.toggle("is-open", isOpen);
  if (libraryAdvancedToggle) {
    libraryAdvancedToggle.classList.toggle("is-open", isOpen);
    libraryAdvancedToggle.setAttribute("aria-expanded", String(isOpen));
  }
}

function updateSendButtonState() {
  if (!nameInput || !fileInput || !sendButton) return;
  const hasName = nameInput.value.trim().length > 0;
  const hasFile = Boolean(fileInput.files && fileInput.files.length);
  sendButton.disabled = !(hasName && hasFile);
}

function updateComposerMetrics() {
  if (!chatForm) return;
  const height = Math.ceil(chatForm.getBoundingClientRect().height || 0);
  if (!height) return;
  document.documentElement.style.setProperty("--composer-height", `${height}px`);
}

function bindComposerMetrics() {
  if (!chatForm) return;
  updateComposerMetrics();
  if (typeof ResizeObserver === "undefined") return;
  if (composerResizeObserver) composerResizeObserver.disconnect();
  composerResizeObserver = new ResizeObserver(() => updateComposerMetrics());
  composerResizeObserver.observe(chatForm);
}

function setToolsDrawerOpen(isOpen) {
  const drawerMode = typeof window.matchMedia === "function" && window.matchMedia("(max-width: 1024px)").matches;
  const nextOpen = Boolean(isOpen) && drawerMode;
  document.body.classList.toggle("tools-open", nextOpen);
  if (toolsOverlay) toolsOverlay.hidden = !nextOpen;
  if (sidebar) sidebar.setAttribute("aria-hidden", String(drawerMode && !nextOpen));
  if (sidebarDrawerToggle) {
    const label = nextOpen ? "Close study tools" : "Open study tools";
    sidebarDrawerToggle.setAttribute("aria-expanded", String(nextOpen));
    sidebarDrawerToggle.setAttribute("aria-label", label);
    sidebarDrawerToggle.setAttribute("title", label);
  }
  if (toolsToggle) {
    toolsToggle.setAttribute("aria-expanded", String(nextOpen));
  }
}

function toggleToolsDrawer() {
  const next = !document.body.classList.contains("tools-open");
  setToolsDrawerOpen(next);
}

function updateAttachmentIndicator() {
  if (!attachStatus) return;
  if (oneShotFile) {
    attachStatus.removeAttribute("title");
    attachStatus.dataset.state = "attached";
    const hideText = oneShotFile.type?.toLowerCase().startsWith("image/");
    const thumbHtml = oneShotPreviewUrl
      ? `<span class="attachment-chip-thumb" style="background-image:url('${escapeHTML(oneShotPreviewUrl)}');"></span>`
      : "";
    attachStatus.innerHTML = `
      <span class="attachment-chip">
        ${thumbHtml}
        <span class="attachment-chip-text">${hideText ? "" : escapeHTML(oneShotFile.name)}</span>
        <button type="button" data-remove="oneshot" aria-label="Remove attachment" title="Remove attachment">✕</button>
      </span>`;
    const btn = attachStatus.querySelector("button[data-remove=\"oneshot\"]");
    if (btn) {
      btn.addEventListener("click", () => {
        if (oneShotPreviewUrl && oneShotPreviewUrl.startsWith("blob:")) URL.revokeObjectURL(oneShotPreviewUrl);
        oneShotFile = null;
        oneShotPreviewUrl = "";
        updateAttachmentIndicator();
      });
    }
    updateResponseContextBar();
    updateComposerMetrics();
    return;
  }
  if (!attachments.length) {
    attachStatus.dataset.state = "empty";
    attachStatus.innerHTML = "";
    attachStatus.removeAttribute("title");
    updateResponseContextBar();
    updateComposerMetrics();
    return;
  }
  attachStatus.removeAttribute("title");
  attachStatus.dataset.state = "attached";
  attachStatus.innerHTML = attachments
    .map((att, idx) => {
      const isImage = isImageAttachment(att);
      const thumb = isImage ? (att.previewUrl || att.url || "") : "";
      const thumbHtml = thumb
        ? `<span class="attachment-chip-thumb" style="background-image:url('${escapeHTML(thumb)}');"></span>`
        : "";
      return `
      <span class="attachment-chip">
        ${thumbHtml}
        <span class="attachment-chip-text">
          ${isImage ? "" : escapeHTML(att.filename)}
        </span>
        <button type="button" data-remove="${idx}" aria-label="Remove attachment" title="Remove attachment">✕</button>
      </span>`;
    })
    .join("");

  attachStatus.querySelectorAll("button[data-remove]").forEach(btn => {
    btn.addEventListener("click", () => {
      const index = Number(btn.getAttribute("data-remove"));
      if (!Number.isNaN(index)) {
        const [removed] = attachments.splice(index, 1);
        if (removed?.fileId) delete attachmentUrlMap[removed.fileId];
        if (removed?.previewUrl && removed.previewUrl.startsWith("blob:")) {
          URL.revokeObjectURL(removed.previewUrl);
        }
        updateAttachmentIndicator();
      }
    });
  });
  updateResponseContextBar();
  updateComposerMetrics();
}

function getAttachmentBadges(att) {
  const badges = [];
  if (typeof att.vectorStoreId === "string" && att.vectorStoreId) {
    badges.push(
      att.vectorStoreStatus === "completed"
        ? { text: "RAG READY", className: "ready" }
        : { text: "RAG INDEXING", className: "warning" },
    );
  } else if (att.vectorStoreId === null) {
    badges.push({ text: "NO RAG", className: "warning" });
  }
  if (att.textKey || att.ocrStatus) {
    if (att.ocrStatus === "error") {
      badges.push({ text: "OCR ERROR", className: "warning", title: att.ocrWarning });
    } else if (att.ocrStatus === "empty") {
      badges.push({ text: "OCR EMPTY", className: "warning", title: att.ocrWarning });
    } else {
      badges.push({ text: "OCR READY", className: "ready" });
    }
  }
  if (att.visionWarning) {
    badges.push({ text: "VISION WARN", className: "warning", title: att.visionWarning });
  }
  return badges;
}

function appendChatMessage(role, text, options = {}) {
  const {
    track = true,
    timestamp,
    model,
    attachments: bubbleAttachments,
    imageUrl,
    imageAlt,
    msgId,
    createdAt,
    docId,
    requestId,
    references,
    evidence,
    extractedKey,
    docTitle,
    sources,
    citations,
    answerSegments,
    renderedMarkdown,
  } = options;
  chatLog.dataset.empty = "false";
  const bubble = document.createElement("div");
  bubble.className = `bubble ${role}`;
  const resolvedMsgId = msgId || generateMessageId();
  bubble.dataset.msgId = resolvedMsgId;
  if (docId) bubble.dataset.docId = docId;
  if (requestId) bubble.dataset.requestId = requestId;
  const renderText = typeof renderedMarkdown === "string" && renderedMarkdown.trim()
    ? renderedMarkdown
    : text;
  if (role === "assistant") {
    const safeSegments = sanitizeAnswerSegmentsForStorage(answerSegments);
    if (safeSegments.length) {
      const citationMap = buildCitationMapFromSegments(safeSegments);
      const normalizedSources = normalizeCitedSources(sources, citationMap);
      renderCitedAnswerSegments(bubble, safeSegments, normalizedSources);
    } else {
      const hasStoredCitations = (Array.isArray(citations) && citations.length) || (Array.isArray(sources) && sources.length);
      const citationMap = buildCitationMapFromStored({ citations, sources });
      const normalizedSources = normalizeCitedSources(sources, citationMap);
      const textBlock = ensureBubbleTextNode(bubble, { reset: true });
      try {
        textBlock.replaceChildren(decorateAssistantResponse(renderText, citationMap, {
          msgId: resolvedMsgId,
          answerText: renderText,
          sources: normalizedSources,
          preserveCitationMap: hasStoredCitations,
        }));
      } catch (err) {
        console.error("[RENDER_ERROR_INIT]", { err, answerLen: (renderText || "").length, msgId: resolvedMsgId });
        const pre = document.createElement("pre");
        pre.style.whiteSpace = "pre-wrap";
        pre.textContent = renderText || "";
        textBlock.replaceChildren(pre);
      }
      enhanceAutoTables(textBlock);
      enhanceTablesIn(textBlock);
      const responseCard = textBlock.firstElementChild;
      if (responseCard && responseCard.classList.contains("response-card")) {
        finalizeResponseCard(responseCard);
      }
      bubble.classList.toggle("has-response", Boolean(responseCard && responseCard.classList.contains("response-card")));
    }
    renderReferenceChips(bubble, references);
    if (evidence && evidence.length) {
      renderEvidenceToggle(bubble, evidence);
    }
  } else {
    const textBlock = ensureBubbleTextNode(bubble, { reset: true });
    textBlock.textContent = text;
  }

  const attachmentsForBubble = cloneMessageAttachments(bubbleAttachments);
  const attachmentNodes = renderBubbleAttachments(attachmentsForBubble);
  if (attachmentNodes) {
    bubble.appendChild(attachmentNodes);
  }

  if (typeof imageUrl === "string" && imageUrl.trim()) {
    const figure = document.createElement("figure");
    figure.className = "bubble-image";
    const img = document.createElement("img");
    img.src = imageUrl;
    img.alt = imageAlt || text || "Generated image";
    img.loading = "lazy";
    figure.appendChild(img);
    if (imageAlt || text) {
      const caption = document.createElement("figcaption");
      caption.textContent = imageAlt || text;
      figure.appendChild(caption);
    }
    bubble.appendChild(figure);
  }

  chatLog.append(bubble);
  chatLog.scrollTop = chatLog.scrollHeight;
  if (track) {
    logMessage(role, text, {
      timestamp,
      model,
      attachments: attachmentsForBubble,
      imageUrl,
      imageAlt: imageAlt || text,
      id: resolvedMsgId,
      createdAt,
      docId,
      requestId,
      references,
      evidence,
      extractedKey,
      docTitle,
      sources,
      citations,
      answerSegments,
      renderedMarkdown,
    });
  }
  return bubble;
}

function clearConversation() {
  startNewConversation();
}

function applyModelSelection(model) {
  const target = modelButtons.find(btn => btn.dataset.model === model);
  if (!target) return;
  modelButtons.forEach(btn => btn.classList.toggle("active", btn === target));
  currentModel = model;
  localStorage.setItem(MODEL_STORAGE_KEY, model);
}

function applyTheme(mode) {
  const isDark = mode === "dark";
  document.body.classList.toggle("theme-dark", isDark);
  document.documentElement.setAttribute("data-theme", isDark ? "dark" : "light");
  if (themeToggle) {
    const label = isDark ? "Switch to light mode" : "Switch to dark mode";
    themeToggle.setAttribute("aria-pressed", String(isDark));
    themeToggle.setAttribute("aria-label", label);
    themeToggle.setAttribute("title", "Toggle theme");
    themeToggle.innerHTML = isDark ? JAGGED_MOON_ICON : LIGHT_BULB_ICON;
  }
}

const savedTheme = localStorage.getItem(THEME_KEY);
applyTheme(savedTheme === "dark" ? "dark" : "light");

function applySidebarCollapsed(isCollapsed) {
  if (!sidebar) return;
  sidebar.classList.toggle("is-collapsed", Boolean(isCollapsed));
  document.body.classList.toggle("sidebar-collapsed", Boolean(isCollapsed));
  if (sidebarCollapseToggle) {
    sidebarCollapseToggle.setAttribute("aria-pressed", String(isCollapsed));
    sidebarCollapseToggle.setAttribute("aria-expanded", String(!isCollapsed));
    const label = isCollapsed ? "Expand sidebar" : "Collapse sidebar";
    sidebarCollapseToggle.setAttribute("aria-label", label);
    sidebarCollapseToggle.setAttribute("title", label);
  }
  safeSetItem(SIDEBAR_COLLAPSE_KEY, String(Boolean(isCollapsed)));
}

function applyMetaPanelVisibility(isVisible) {
  if (!metaDataPanel) return;
  metaDataPanel.hidden = !isVisible;
  metaDataPanel.setAttribute("aria-hidden", String(!isVisible));
  if (metaPanelToggle) {
    metaPanelToggle.setAttribute("aria-pressed", String(isVisible));
    metaPanelToggle.textContent = isVisible ? "Hide Debug" : "Show Debug";
  }
  safeSetItem(META_PANEL_VISIBLE_KEY, String(Boolean(isVisible)));
}

const META_UNLOCK_KEY = "owen.metaDataUnlocked";
const HISTORY_COLLAPSE_KEY = "owen.chatHistoryCollapsed";
const LIBRARY_COLLAPSE_KEY = "owen.libraryPanelCollapsed";
const RETRIEVE_COLLAPSE_KEY = "owen.retrievePanelCollapsed";
const MACHINE_COLLAPSE_KEY = "owen.machinePanelCollapsed";
const MACHINE_UNLOCK_KEY = "owen.machineUnlocked";
const MACHINE_UPLOAD_COLLAPSE_KEY = "owen.machineUploadCollapsed";
const MACHINE_TXT_COLLAPSE_KEY = "owen.machineTxtCollapsed";
const MACHINE_STUDY_COLLAPSE_KEY = "owen.machineStudyCollapsed";

document.addEventListener("click", (e) => {
  const target = e.target instanceof Element ? e.target : null;
  const btn = target ? target.closest("button.subpanel-toggle") : null;
  if (!btn || btn.dataset.subpanelBound === "true") return;
  const panel = btn.closest(".machine-subpanel");
  if (!panel) return;
  const body = panel.querySelector(".machine-subpanel__body");
  if (!body) return;
  e.preventDefault();
  e.stopPropagation();
  const collapsed = !panel.classList.contains("is-collapsed");
  panel.classList.toggle("is-collapsed", collapsed);
  if (body instanceof HTMLElement) {
    body.hidden = collapsed;
  }
  btn.setAttribute("aria-expanded", String(!collapsed));
  const storageKey = panel.dataset.storageKey || "";
  if (storageKey) {
    try {
      localStorage.setItem(storageKey, String(collapsed));
    } catch {}
  }
}, true);

if (metaDataPanel && metaUnlockInput && metaLockBtn) {
  const unlocked = localStorage.getItem(META_UNLOCK_KEY) === "true";
  applyMetaUnlockState(unlocked);
  const tryUnlock = () => {
    if (metaUnlockInput.value === "1234") {
      applyMetaUnlockState(true);
      metaUnlockInput.value = "";
      setMetaUnlockError("");
      return;
    }
    if (metaUnlockInput.value) {
      setMetaUnlockError("Wrong password");
      metaUnlockInput.value = "";
    }
  };
  metaUnlockInput.addEventListener("input", () => setMetaUnlockError(""));
  metaUnlockInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      tryUnlock();
    }
  });
  metaLockBtn.addEventListener("click", () => {
    if (!metaDataUnlocked) {
      tryUnlock();
      return;
    }
    applyMetaUnlockState(false);
    metaUnlockInput.value = "";
  });
}

if (metaToggle) {
  metaToggle.addEventListener("click", () => {
    if (!metaDataUnlocked) return;
    applyMetaExpandedState(!metaDataExpanded);
  });
}

if (historyPanel && historyToggle) {
  const collapsed = localStorage.getItem(HISTORY_COLLAPSE_KEY) === "true";
  if (collapsed) historyPanel.classList.add("is-collapsed");
  historyToggle.addEventListener("click", () => {
    const next = !historyPanel.classList.contains("is-collapsed");
    historyPanel.classList.toggle("is-collapsed", next);
    try {
      localStorage.setItem(HISTORY_COLLAPSE_KEY, String(next));
    } catch {}
  });
}

if (libraryPanel && libraryToggle && libraryPanelBody) {
  const collapsed = localStorage.getItem(LIBRARY_COLLAPSE_KEY) === "true";
  if (collapsed) libraryPanel.classList.add("is-collapsed");
  libraryToggle.addEventListener("click", () => {
    const next = !libraryPanel.classList.contains("is-collapsed");
    libraryPanel.classList.toggle("is-collapsed", next);
    try {
      localStorage.setItem(LIBRARY_COLLAPSE_KEY, String(next));
    } catch {}
  });
}

if (retrievePanel && retrieveToggle && retrievePanelBody) {
  const collapsed = localStorage.getItem(RETRIEVE_COLLAPSE_KEY) === "true";
  if (collapsed) retrievePanel.classList.add("is-collapsed");
  retrieveToggle.addEventListener("click", () => {
    const next = !retrievePanel.classList.contains("is-collapsed");
    retrievePanel.classList.toggle("is-collapsed", next);
    try {
      localStorage.setItem(RETRIEVE_COLLAPSE_KEY, String(next));
    } catch {}
  });
}

if (machinePanel && machineToggle && machinePanelBody) {
  const collapsed = localStorage.getItem(MACHINE_COLLAPSE_KEY) === "true";
  const unlocked = localStorage.getItem(MACHINE_UNLOCK_KEY) === "true";
  if (collapsed) machinePanel.classList.add("is-collapsed");
  applyMachineUnlockState(unlocked);
  machineToggle.addEventListener("click", () => {
    if (!machineUnlocked) {
      machineUnlockInput?.focus();
      return;
    }
    const next = !machinePanel.classList.contains("is-collapsed");
    machinePanel.classList.toggle("is-collapsed", next);
    try {
      localStorage.setItem(MACHINE_COLLAPSE_KEY, String(next));
    } catch {}
  });
  if (machineUnlockInput) {
    const tryUnlock = () => {
      if (machineUnlockInput.value === "1234") {
        applyMachineUnlockState(true);
        return;
      }
      if (machineUnlockInput.value) {
        flashMachineUnlockError();
        machineUnlockInput.value = "";
      }
    };
    machineUnlockInput.addEventListener("input", () => {
      machineUnlockInput.classList.remove("is-error");
      if (machineUnlockErrorTimer) {
        clearTimeout(machineUnlockErrorTimer);
        machineUnlockErrorTimer = null;
      }
    });
    machineUnlockInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        tryUnlock();
      }
    });
  }
  if (machineLockBtn) {
    machineLockBtn.addEventListener("click", () => {
      if (!machineUnlocked) {
        if (machineUnlockInput) {
          machineUnlockInput.focus();
        }
        const inputValue = machineUnlockInput?.value || "";
        if (inputValue) {
          if (inputValue === "1234") {
            applyMachineUnlockState(true);
          } else {
            flashMachineUnlockError();
            if (machineUnlockInput) machineUnlockInput.value = "";
          }
        }
        return;
      }
      applyMachineUnlockState(false);
    });
  }
}

if (storedOcrSession && storedOcrSession.fileHash) {
  activePdfSession = {
    file: null,
    arrayBuffer: null,
    fileHash: storedOcrSession.fileHash,
    extractedKey: storedOcrSession.extractedKey || null,
    pageCount: storedOcrSession.pageCount || null,
    nextPageStart: storedOcrSession.processedRanges?.length
      ? storedOcrSession.processedRanges[storedOcrSession.processedRanges.length - 1].end + 1
      : 1,
    lastQuestion: "",
    filename: storedOcrSession.filename,
    backgroundPromise: null,
    processedRanges: storedOcrSession.processedRanges || [],
  };
}

localStorage.setItem(MODEL_STORAGE_KEY, currentModel);
applyModelSelection(currentModel);

async function generatePreviewDataUrl(file) {
  try {
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : "");
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
    if (!dataUrl) return "";
    const img = new Image();
    const loadImg = new Promise((resolve, reject) => {
      img.onload = () => resolve(true);
      img.onerror = reject;
    });
    img.src = dataUrl;
    await loadImg;
    const maxSide = 320;
    const scale = Math.min(1, maxSide / Math.max(img.width || maxSide, img.height || maxSide));
    const targetW = Math.max(1, Math.round((img.width || maxSide) * scale));
    const targetH = Math.max(1, Math.round((img.height || maxSide) * scale));
    const canvas = document.createElement("canvas");
    canvas.width = targetW;
    canvas.height = targetH;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0, targetW, targetH);
    return canvas.toDataURL("image/png", 0.82);
  } catch {
    return "";
  }
}

runRenderFallbackSelfTest();
runEmptyFinalRegressionTest();

async function uploadFile() {
  const subject = nameInput.value.trim();
  const file = fileInput.files?.[0] ?? null;
  if (!subject) {
    setUploadStatusState("error", "Please enter a name before uploading.");
    nameInput.focus();
    return;
  }
  if (!file) {
    setUploadStatusState("error", "Choose a file to upload.");
    fileButton.focus();
    return;
  }

  sendButton.disabled = true;
  sendButton.textContent = "Sending...";
  fileButton.disabled = true;
  setUploadStatusState("pending", "Uploading to R2 + mirroring to OpenAI Files...");

  try {
    const payload = await uploadToWorker(file, subject);

    if (payload.file_id) {
      const displayName = payload.filename || file.name;
      const downloadUrl = `/api/file?bucket=${encodeURIComponent(payload.bucket)}&key=${encodeURIComponent(payload.key)}&filename=${encodeURIComponent(displayName)}`;
      const vectorStatus = payload.vector_store_status || (payload.vector_store_id ? "processing" : undefined);
      attachments.push({
        filename: displayName,
        fileId: payload.file_id,
        visionFileId: payload.vision_file_id || payload.visionFileId,
        bucket: payload.bucket,
        key: payload.key,
        url: downloadUrl,
        vectorStoreId: payload.vector_store_id || null,
        vectorStoreStatus: vectorStatus,
        textKey: payload.ocr_text_key || null,
        ocrStatus: payload.ocr_status || (payload.ocr_text_key ? "ready" : null),
        ocrWarning: payload.ocr_warning || null,
        visionWarning: payload.vision_warning || null,
      });
      attachmentUrlMap[payload.file_id] = downloadUrl;
      summarizeAttachment(payload.file_id, displayName);
      updateAttachmentIndicator();

      const ragReady = vectorStatus === "completed";
      let statusState = ragReady ? "success" : "warning";
      if (!payload.vector_store_id) statusState = "warning";
      if (payload.ocr_status === "error") statusState = "warning";
      if (payload.vector_warning) statusState = "warning";
      renderUploadSuccess(payload, displayName, statusState);
      setPipelineLastRun(1);
      loadLibraryList().catch(() => {});
    } else {
      setUploadStatusState("error", "Uploaded to R2 but failed to mirror to OpenAI. Attach via the + button to retry.");
    }
    fileInput.value = "";
    fileNameDisplay.textContent = "No file selected";
    updateSendButtonState();
  } catch (error) {
    console.error(error);
    setUploadStatusState("error", error instanceof Error ? error.message : "Unexpected upload error.");
  } finally {
    sendButton.textContent = "Upload lecture";
    fileButton.disabled = false;
    updateSendButtonState();
  }
}

async function retrieveFileFromR2({ key, bucket, autoDownload = true, statusEl = retrieveStatus, emitChat = true }) {
  const bucketParam = bucket && bucket.trim() ? bucket.trim() : "";
  const url = `/api/file?key=${encodeURIComponent(key)}${bucketParam ? `&bucket=${encodeURIComponent(bucketParam)}` : ""}`;
  setStatus(statusEl, "pending", `Fetching ${key} ${bucketParam ? `from ${bucketParam}` : "from all buckets"}...`);
  try {
    const res = await fetch(url);
    if (!res.ok) {
      const msg = await res.text().catch(() => "");
      throw new Error(msg || `File not found for key "${key}".`);
    }
    const disposition = res.headers.get("content-disposition") || "";
    const match = disposition.match(/filename=\"?([^\";]+)\"?/i);
    const filename = match?.[1] || key.split("/").pop() || "download";
    const blob = await res.blob();
    const objectUrl = URL.createObjectURL(blob);

    if (autoDownload) {
      const a = document.createElement("a");
      a.href = objectUrl;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(objectUrl), 1500);
    } else {
      URL.revokeObjectURL(objectUrl);
    }

    const downloadUrl = `${url}&filename=${encodeURIComponent(filename)}`;
    if (emitChat) {
      const attachmentsForBubble = [{ filename, url: downloadUrl, bucket: bucketParam || undefined, key }];
      const originLabel = bucketParam ? `bucket ${bucketParam}` : "any bucket";
      const bubble = appendChatMessage("assistant", `Here’s your file from ${originLabel}: ${filename}`, {
        attachments: attachmentsForBubble,
      });
      renderMathIfReady(bubble);
      logMessage("assistant", `File ready: ${filename}`, { attachments: attachmentsForBubble });
    }

    setStatus(statusEl, "success", `Ready: ${filename}`);
    return { filename, downloadUrl };
  } catch (error) {
    console.error(error);
    setStatus(statusEl, "error", error instanceof Error ? error.message : "Unable to retrieve file.");
    throw error;
  }
}

function normalizeAccessCodeInput(value) {
  return (value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

const ACCESS_CODE_REGEX = /^[A-HJ-NP-Z2-9]{6,10}$/;

function renderRetrieveStudyGuide(manifest) {
  if (!retrieveResult) return;
  if (!manifest) {
    retrieveResult.innerHTML = "";
    return;
  }
  const title = manifest.title || "Study Guide";
  const assets = Array.isArray(manifest.assets) ? manifest.assets : [];
  const studyGuide = assets.find(asset => typeof asset.kind === "string" && asset.kind.toLowerCase().includes("study"));
  const anki = assets.find(asset => typeof asset.kind === "string" && asset.kind.toLowerCase().includes("anki"));
  const fallbackLinks = assets.filter(asset => asset && asset.downloadUrl && asset !== studyGuide && asset !== anki);
  const actions = [];
  if (studyGuide?.downloadUrl) {
    const guideFilename = studyGuide.filename || "study-guide.html";
    actions.push(
      `<a class="btn btn-primary" href="${escapeHTML(studyGuide.downloadUrl)}" download="${escapeHTML(guideFilename)}">Download Study Guide</a>`
    );
  }
  if (anki?.downloadUrl) {
    const ankiFilename = anki.filename || "anki.apkg";
    actions.push(
      `<a class="btn btn-ghost" href="${escapeHTML(anki.downloadUrl)}" download="${escapeHTML(ankiFilename)}">Download Anki</a>`
    );
  }
  fallbackLinks.forEach((asset) => {
    actions.push(
      `<a class="btn btn-ghost" href="${escapeHTML(asset.downloadUrl)}" download="${escapeHTML(asset.filename || "download")}">Download ${escapeHTML(asset.kind || "Asset")}</a>`
    );
  });
  retrieveResult.innerHTML = `
    <div class="retrieve-title">${escapeHTML(title)}</div>
    <div class="retrieve-actions">
      ${actions.join("")}
    </div>
  `.trim();
}

async function retrieveStudyGuideByCode() {
  const raw = retrieveInput?.value || "";
  const code = normalizeAccessCodeInput(raw);
  if (retrieveInput) retrieveInput.value = code;
  if (!code) {
    setStatus(retrieveStatus, "error", "Enter a student access code first.");
    retrieveInput?.focus();
    renderRetrieveStudyGuide(null);
    return;
  }
  if (!ACCESS_CODE_REGEX.test(code)) {
    setStatus(retrieveStatus, "error", "Access code should be 6-10 characters (A-Z, 2-9).");
    retrieveInput?.focus();
    renderRetrieveStudyGuide(null);
    return;
  }
  setStatus(retrieveStatus, "pending", "Looking up code...");
  renderRetrieveStudyGuide(null);
  try {
    const resp = await fetch(`/api/retrieve/study-guide?code=${encodeURIComponent(code)}`);
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok || payload?.error) {
      const msg = payload?.error || payload?.message || "Code not found.";
      throw new Error(msg);
    }
    renderRetrieveStudyGuide(payload);
    setStatus(retrieveStatus, "success", "Study guide ready.");
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to retrieve study guide.";
    setStatus(retrieveStatus, "error", message);
    renderRetrieveStudyGuide(null);
  }
}

function renderRetrieveAnki(result) {
  if (!retrieveResult) return;
  if (!result) {
    retrieveResult.innerHTML = "";
    return;
  }
  const filename = result.filename || "cards.tsv";
  const downloadUrl = result.downloadUrl || "";
  const actions = downloadUrl
    ? `<a class="btn btn-primary" href="${escapeHTML(downloadUrl)}" download="${escapeHTML(filename)}">Download cards.tsv</a>`
    : "";
  retrieveResult.innerHTML = `
    <div class="retrieve-title">Anki Deck</div>
    <div class="retrieve-actions">
      ${actions}
    </div>
  `.trim();
}

function extractFilenameFromDisposition(header) {
  if (!header) return "";
  const match = header.match(/filename\*?=(?:UTF-8''|\"?)([^\";]+)/i);
  if (!match) return "";
  try {
    return decodeURIComponent(match[1]);
  } catch {
    return match[1];
  }
}

async function retrieveAnkiByCode() {
  const raw = retrieveInput?.value || "";
  const code = normalizeAccessCodeInput(raw);
  if (retrieveInput) retrieveInput.value = code;
  if (!code) {
    setStatus(retrieveStatus, "error", "Enter a student access code first.");
    retrieveInput?.focus();
    renderRetrieveAnki(null);
    return;
  }
  if (!ACCESS_CODE_REGEX.test(code)) {
    setStatus(retrieveStatus, "error", "Access code should be 6-10 characters (A-Z, 2-9).");
    retrieveInput?.focus();
    renderRetrieveAnki(null);
    return;
  }
  setStatus(retrieveStatus, "pending", "Looking up code...");
  renderRetrieveAnki(null);
  try {
    const resp = await fetch(`/api/anki/download?code=${encodeURIComponent(code)}`);
    if (!resp.ok) {
      const payload = await resp.json().catch(() => ({}));
      const msg = payload?.error || payload?.message || "Invalid code.";
      throw new Error(msg);
    }
    const disposition = resp.headers.get("content-disposition") || "";
    const filename = extractFilenameFromDisposition(disposition) || "cards.tsv";
    const blob = await resp.blob();
    const downloadUrl = URL.createObjectURL(blob);
    renderRetrieveAnki({ downloadUrl, filename });
    setStatus(retrieveStatus, "success", "Anki deck ready.");
    const link = retrieveResult?.querySelector("a");
    if (link && link instanceof HTMLAnchorElement) {
      link.click();
    }
    setTimeout(() => {
      URL.revokeObjectURL(downloadUrl);
    }, 60_000);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to retrieve Anki deck.";
    setStatus(retrieveStatus, "error", message);
    renderRetrieveAnki(null);
  }
}

function retrieveByCode() {
  const type = retrieveTypeSelect?.value || "study-guide";
  if (type === "anki") {
    return retrieveAnkiByCode();
  }
  return retrieveStudyGuideByCode();
}

async function uploadChatAttachment(file) {
  if (!attachStatus) return;
  attachStatus.dataset.state = "attached";
  attachStatus.innerHTML = `
    <span class="attachment-chip loading">
      <span>Attaching ${escapeHTML(file.name)}</span>
      <span class="loading-track"></span>
    </span>`;
  const task = (async () => {
    let previewUrl = "";
    let mimeType = "";
    try {
      mimeType = typeof file.type === "string" ? file.type : "";
      previewUrl = mimeType.startsWith("image/") ? URL.createObjectURL(file) : "";
      const payload = await uploadToWorker(file, "chat-attachment");
      const displayName = payload.filename || file.name;
      const downloadUrl = `/api/file?bucket=${encodeURIComponent(payload.bucket)}&key=${encodeURIComponent(payload.key)}&filename=${encodeURIComponent(displayName)}`;
      const vectorStatus = payload.vector_store_status || (payload.vector_store_id ? "processing" : undefined);
      attachments.push({
        filename: displayName,
        fileId: payload.file_id,
        visionFileId: payload.vision_file_id || payload.visionFileId,
        bucket: payload.bucket,
        key: payload.key,
        url: downloadUrl,
        mimeType,
        previewUrl,
        vectorStoreId: payload.vector_store_id || null,
        vectorStoreStatus: vectorStatus,
        textKey: payload.ocr_text_key || null,
        ocrStatus: payload.ocr_status || (payload.ocr_text_key ? "ready" : null),
        ocrWarning: payload.ocr_warning || null,
        visionWarning: payload.vision_warning || null,
      });
      attachmentUrlMap[payload.file_id] = downloadUrl;
      summarizeAttachment(payload.file_id, displayName);
      updateAttachmentIndicator();
      if (payload.ocr_warning) {
        attachStatus.setAttribute("title", payload.ocr_warning);
      }
    } catch (error) {
      console.error(error);
      attachStatus.dataset.state = "error";
      attachStatus.textContent = error instanceof Error ? error.message : "Could not attach file.";
      if (previewUrl && previewUrl.startsWith("blob:")) {
        URL.revokeObjectURL(previewUrl);
      }
      throw error;
    }
  })();
  attachmentUploads.add(task);
  try {
    await task;
  } finally {
    attachmentUploads.delete(task);
  }
}

async function uploadToWorker(file, name) {
  const formData = new FormData();
  formData.append("file", file);
  const destination = bucketSelect.value || DESTINATIONS[0].value;
  formData.append("bucket", destination);
  formData.append("destination", destination);
  if (name) formData.append("name", name);

  const response = await fetch("/api/upload", { method: "POST", body: formData });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !payload?.ok) {
    throw new Error(payload?.error || "Upload failed. Try again.");
  }
  if (!payload.file_id) {
    throw new Error("Attachment could not be mirrored to OpenAI. Please retry.");
  }
  return payload;
}

function setThinkingStatus(bubble, text) {
  if (!bubble) return;
  const status = ensureBubbleStatusNode(bubble);
  const label = status?.querySelector(".bubble-status-text");
  if (label) label.textContent = text;
  chatLog.scrollTop = chatLog.scrollHeight;
}

async function ingestPdfToWorker(file, arrayBuffer, fileHash, pageCap = OCR_PAGE_LIMIT, signal) {
  throwIfAborted(signal);
  const form = new FormData();
  form.append("file", file);
  form.append("fileHash", fileHash);
  form.append("maxPages", String(pageCap));
  const resp = await fetch("/api/pdf-ingest", { method: "POST", body: form, signal });
  const payload = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const msg = payload?.error || payload?.details || "PDF ingest failed.";
    throw new Error(msg);
  }
  return payload;
}

async function askDocWithRetrieval(question, extractedKey, fileHash, signal) {
  throwIfAborted(signal);
  const resp = await fetch("/api/ask-doc", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ question, extractedKey, fileHash }),
    signal,
  });
  const payload = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const msg = payload?.error || payload?.details || "Document Q&A failed.";
    throw new Error(msg);
  }
  return payload?.answer || "(no answer returned)";
}

async function ocrPageImage({ fileHash, fileId, pageIndex, blob, filename, signal }) {
  let attempt = 0;
  let lastError = "";
  while (attempt < OCR_MAX_RETRIES) {
    throwIfAborted(signal);
    const fd = new FormData();
    if (fileHash) fd.append("fileHash", fileHash);
    if (fileId) fd.append("fileId", fileId);
    fd.append("pageIndex", String(pageIndex));
    fd.append("image", blob, filename || `page-${pageIndex + 1}.jpg`);
    const resp = await fetch("/api/ocr-page", { method: "POST", body: fd, signal });
    const payload = await resp.json().catch(() => ({}));
    if (resp.ok) {
      const text = typeof payload.text === "string" ? payload.text : "";
      if (!text.trim()) {
        throw new Error(`OCR returned no text for page ${pageIndex + 1}.`);
      }
      return { pageIndex, text };
    }
    lastError = payload?.error || payload?.details || resp.statusText || "OCR page request failed.";
    if (resp.status === 429 || resp.status === 503) {
      const delayMs = Math.min(4000, Math.pow(2, attempt) * 600 + Math.random() * 300);
      console.warn("[OCR] retry", { pageIndex, attempt: attempt + 1, status: resp.status, delayMs });
      await sleep(delayMs);
      attempt += 1;
      continue;
    }
    throw new Error(lastError);
  }
  throw new Error(lastError || `OCR failed for page ${pageIndex + 1}.`);
}

async function finalizeOcrText({ fileHash, fileId, pages, filename, totalPages, signal }) {
  throwIfAborted(signal);
  const resp = await fetch("/api/ocr-finalize", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ fileHash, fileId, pages, filename, totalPages }),
    signal,
  });
  const payload = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const msg = payload?.error || payload?.details || "OCR finalize failed.";
    throw new Error(msg);
  }
  if (!payload?.extractedKey) {
    throw new Error("OCR finalize did not return an extractedKey.");
  }
  return payload;
}

async function runPdfOcrFlow({ file, prompt, pageCap = OCR_PAGE_LIMIT, thinkingBubble, startPage = 1, signal }) {
  throwIfAborted(signal);
  const wantsAllPages = wantsEntireDocument(prompt) || fullDocumentMode;
  const requestedCap = wantsAllPages ? Math.max(pageCap, 9999) : pageCap;
  setThinkingStatus(thinkingBubble, "Hashing PDF…");
  const arrayBuffer = await file.arrayBuffer();
  throwIfAborted(signal);
  const fileHash = await computeSha256(arrayBuffer);
  throwIfAborted(signal);
  setThinkingStatus(thinkingBubble, "Checking PDF…");
  const ingest = await ingestPdfToWorker(file, arrayBuffer, fileHash, requestedCap, signal);
  throwIfAborted(signal);
  const pageCount = ingest?.pageCount || 0;
  console.log("[PDF] ingest result", {
    method: ingest?.method,
    extractionStatus: ingest?.extractionStatus,
    pageCount,
    fileHash,
  });
  activePdfSession = {
    file,
    arrayBuffer,
    fileHash,
    extractedKey: ingest?.extractedKey || null,
    pageCount,
    nextPageStart: startPage,
    lastQuestion: prompt,
    filename: file.name,
    backgroundPromise: null,
    processedRanges: ingest?.manifest?.ranges || [],
  };
  persistOcrSession(activePdfSession);

  if (ingest?.extractionStatus === "ok" && ingest?.extractedKey) {
    const method = ingest?.method || "embedded";
    setThinkingStatus(thinkingBubble, method === "cache" ? "Using cached text…" : "Using embedded text…");
    const answer = await askDocWithRetrieval(prompt, ingest.extractedKey, fileHash, signal);
    throwIfAborted(signal);
    let continueRange = null;
    if (
      !wantsAllPages &&
      ingest?.manifest?.pageCount &&
      ingest?.manifest?.pagesProcessed &&
      ingest.manifest.pageCount > ingest.manifest.pagesProcessed
    ) {
      const nextStart = ingest.manifest.pagesProcessed + 1;
      continueRange = {
        start: nextStart,
        end: Math.min(ingest.manifest.pageCount, nextStart + OCR_PAGE_LIMIT - 1),
        totalPages: ingest.manifest.pageCount,
      };
    }
    activePdfSession.nextPageStart = continueRange?.start || pageCount + 1;
    activePdfSession.extractedKey = ingest.extractedKey;
    persistOcrSession(activePdfSession);
    return { answer, warning: "", continueRange, extractedKey: ingest.extractedKey };
  }

  setThinkingStatus(thinkingBubble, "Rendering pages for OCR…");
  const pdfjsLib = await loadPdfJsLib();
  throwIfAborted(signal);
  const pdfDoc = await pdfjsLib.getDocument({
    data: arrayBuffer,
    disableWorker: true,
  }).promise;
  throwIfAborted(signal);
  const totalPages = pageCount || pdfDoc?.numPages || requestedCap;
  const start = Math.max(1, startPage);
  const capForRange = wantsAllPages ? totalPages : pageCap;
  const end = Math.min(totalPages, start + capForRange - 1);
  const pageNumbers = [];
  setThinkingStatus(thinkingBubble, `Processing pages ${start}-${end} of ${totalPages}`);
  for (let i = start; i <= end; i += 1) pageNumbers.push(i);

  let resolvePartial;
  let rejectPartial;
  const partialPromise = new Promise((resolve, reject) => {
    resolvePartial = resolve;
    rejectPartial = reject;
  });
  const state = {
    allPages: [],
    partialPages: [],
    extractedKey: ingest?.extractedKey || null,
    answer: "",
    resolved: false,
    finalizing: false,
    processedRanges: [...(activePdfSession?.processedRanges || [])],
  };

  const finalizeAndAnswer = async (pagesSubset) => {
    state.finalizing = true;
    try {
      const ordered = [...pagesSubset].sort((a, b) => a.pageIndex - b.pageIndex);
      const finalize = await finalizeOcrText({
        fileHash,
        pages: ordered,
        filename: file.name,
        totalPages,
        signal,
      });
      state.extractedKey = finalize.extractedKey;
      state.partialPages = ordered;
      state.answer = await askDocWithRetrieval(prompt, finalize.extractedKey, fileHash, signal);
      state.resolved = true;
      const newRange = { start, end };
      state.processedRanges.push(newRange);
      activePdfSession.processedRanges = state.processedRanges;
      activePdfSession.extractedKey = finalize.extractedKey;
      persistOcrSession(activePdfSession);
      resolvePartial();
    } catch (err) {
      rejectPartial(err);
      throw err;
    } finally {
      state.finalizing = false;
    }
  };

  let pendingBatch = [];
  const poolPromise = promisePool(pageNumbers, OCR_CONCURRENCY, async (pageNumber) => {
    if (activePdfSession?.cancelled) {
      throw createAbortError("Processing cancelled.");
    }
    throwIfAborted(signal);
    if (thinkingBubble && thinkingBubble.isConnected) {
      setThinkingStatus(thinkingBubble, `OCR page ${state.allPages.length + 1} / ${pageNumbers.length}`);
    }
    const render = await renderPdfPageToBlob(pdfjsLib, pdfDoc, pageNumber);
    throwIfAborted(signal);
    const pageResult = await ocrPageImage({
      fileHash,
      fileId: ingest?.fileId || null,
      pageIndex: pageNumber - 1,
      blob: render.blob,
      filename: file.name,
      signal,
    });
    state.allPages.push(pageResult);
    pendingBatch.push(pageResult);
    const shouldFlush = pendingBatch.length >= OCR_SAVE_BATCH || pageNumber === end;
    if (shouldFlush && pendingBatch.length) {
      await finalizeOcrText({
        fileHash,
        pages: [...pendingBatch],
        filename: file.name,
        totalPages,
        signal,
      });
      pendingBatch = [];
    }
    if (!state.resolved && !state.finalizing && state.allPages.length >= Math.min(EARLY_OCR_READY_COUNT, pageNumbers.length)) {
      await finalizeAndAnswer(state.allPages);
    }
  });

  activePdfSession.backgroundPromise = poolPromise;
  poolPromise
    .then(async () => {
      if (!state.resolved) {
        await finalizeAndAnswer(state.allPages);
      }
      const partialSet = new Set(state.partialPages.map(p => p.pageIndex));
      const remaining = state.allPages
        .filter(page => !partialSet.has(page.pageIndex))
        .sort((a, b) => a.pageIndex - b.pageIndex);
      if (remaining.length) {
        const finalize = await finalizeOcrText({
          fileHash,
          pages: remaining,
          filename: file.name,
          totalPages,
          signal,
        });
        state.extractedKey = finalize.extractedKey || state.extractedKey;
      }
      pdfDoc.cleanup?.();
      pdfDoc.destroy?.();
      activePdfSession = {
        file,
        arrayBuffer,
        fileHash,
        extractedKey: state.extractedKey,
        pageCount: totalPages,
        nextPageStart: end + 1,
        lastQuestion: prompt,
        filename: file.name,
        backgroundPromise: null,
        processedRanges: state.processedRanges,
      };
      persistOcrSession(activePdfSession);
    })
    .catch(err => {
      if (!isAbortError(err)) {
        console.error("[OCR] background failure", err);
      }
      if (activePdfSession) {
        activePdfSession.backgroundPromise = null;
      }
      pdfDoc.cleanup?.();
      pdfDoc.destroy?.();
      if (!state.resolved) {
        rejectPartial(err);
      }
    });

  await partialPromise;
  throwIfAborted(signal);
  const continueRange = !wantsAllPages && totalPages > end
    ? { start: end + 1, end: Math.min(totalPages, end + OCR_PAGE_LIMIT), totalPages }
    : null;
  const warning = `Processed pages ${start}-${end}.`;
  return { answer: state.answer, warning, continueRange, extractedKey: state.extractedKey, fileHash };
}

async function continuePdfRange({ start, end, prompt, thinkingBubble, signal }) {
  throwIfAborted(signal);
  if (!activePdfSession || !activePdfSession.arrayBuffer || !activePdfSession.fileHash) {
    throw new Error("PDF session expired. Re-upload the same PDF to continue OCR.");
  }
  if (activePdfSession.backgroundPromise) {
    try {
      await activePdfSession.backgroundPromise;
    } catch {
      // already logged upstream
    } finally {
      activePdfSession.backgroundPromise = null;
    }
  }
  throwIfAborted(signal);
  const pdfjsLib = await loadPdfJsLib();
  let pdfDoc = null;
  try {
    pdfDoc = await pdfjsLib.getDocument({
      data: activePdfSession.arrayBuffer,
      disableWorker: true,
    }).promise;
    throwIfAborted(signal);
    const totalPages = activePdfSession.pageCount || pdfDoc?.numPages || end;
    const safeEnd = Math.min(end, totalPages);
    const pageNumbers = [];
    for (let i = start; i <= safeEnd; i += 1) pageNumbers.push(i);
    setThinkingStatus(thinkingBubble, `Processing pages ${start}-${safeEnd} of ${totalPages}`);
    const results = [];
    await promisePool(pageNumbers, OCR_CONCURRENCY, async (pageNumber) => {
      if (activePdfSession?.cancelled) {
        throw createAbortError("Processing cancelled.");
      }
      throwIfAborted(signal);
      if (thinkingBubble && thinkingBubble.isConnected) {
        setThinkingStatus(thinkingBubble, `OCR page ${results.length + 1} / ${pageNumbers.length}`);
      }
      const render = await renderPdfPageToBlob(pdfjsLib, pdfDoc, pageNumber);
      throwIfAborted(signal);
      const pageResult = await ocrPageImage({
        fileHash: activePdfSession.fileHash,
        fileId: null,
        pageIndex: pageNumber - 1,
        blob: render.blob,
        filename: activePdfSession.filename,
        signal,
      });
      results.push(pageResult);
    });
    results.sort((a, b) => a.pageIndex - b.pageIndex);
    const finalize = await finalizeOcrText({
      fileHash: activePdfSession.fileHash,
      pages: results,
      filename: activePdfSession.filename,
      totalPages,
      signal,
    });
    console.log("[PDF] Continue OCR stored", {
      start,
      end: safeEnd,
      totalPages,
      fileHash: activePdfSession.fileHash,
    });
    activePdfSession.extractedKey = finalize.extractedKey || activePdfSession.extractedKey;
    activePdfSession.nextPageStart = safeEnd + 1;
    activePdfSession.lastQuestion = prompt || activePdfSession.lastQuestion;
    activePdfSession.processedRanges = [
      ...(activePdfSession.processedRanges || []),
      { start, end: safeEnd },
    ];
    persistOcrSession(activePdfSession);

    const answer = await askDocWithRetrieval(
      prompt || activePdfSession.lastQuestion || "Summarize the new pages",
      finalize.extractedKey,
      activePdfSession.fileHash,
      signal,
    );
    throwIfAborted(signal);
    const continueRange = safeEnd < totalPages
      ? { start: safeEnd + 1, end: Math.min(totalPages, safeEnd + OCR_PAGE_LIMIT), totalPages }
      : null;
    const warning = `Processed pages ${start}-${safeEnd}.`;
    return { answer, warning, continueRange };
  } finally {
    pdfDoc?.cleanup?.();
    pdfDoc?.destroy?.();
  }
}

function renderContinuePrompt(range, prompt) {
  const prevStart = Math.max(1, range.start - OCR_PAGE_LIMIT);
  const prevEnd = Math.max(prevStart, range.start - 1);
  const label = `Processed pages ${prevStart}-${prevEnd}. Continue pages ${range.start}-${range.end}?`;
  const bubble = appendChatMessage("assistant", label, { track: false });
  const btn = document.createElement("button");
  btn.type = "button";
  btn.textContent = `Continue ${range.start}-${range.end}`;
  btn.className = "continue-btn";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.textContent = "Cancel";
  cancelBtn.className = "continue-cancel-btn";
  bubble.appendChild(btn);
  bubble.appendChild(cancelBtn);
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    cancelBtn.disabled = true;
    btn.textContent = "Continuing…";
    const thinking = showThinkingBubble();
    const controller = registerOwenAbortController(new AbortController(), thinking);
    const { signal } = controller;
    try {
      const result = await continuePdfRange({
        start: range.start,
        end: range.end,
        prompt,
        thinkingBubble: thinking,
        signal,
      });
      if (thinking && thinking.isConnected) thinking.remove();
      const finalText = result.warning ? `${result.answer}\n\n_${result.warning}_` : result.answer;
      appendChatMessage("assistant", finalText, {
        model: "gpt-5-mini",
        docId: activePdfSession?.fileHash || null,
        extractedKey: activePdfSession?.extractedKey || result?.extractedKey || null,
      });
      if (result?.continueRange) {
        renderContinuePrompt(result.continueRange, prompt);
      }
    } catch (err) {
      if (isAbortError(err)) {
        if (thinking && thinking.isConnected) clearThinkingBubble(thinking);
        return;
      }
      console.error(err);
      if (thinking && thinking.isConnected) thinking.remove();
      const errMsg = err instanceof Error ? err.message : "Continue OCR failed.";
      const errBubble = appendChatMessage("assistant", errMsg, { track: false });
      errBubble.classList.add("error");
    } finally {
      clearOwenAbortController(controller);
    }
  });
  cancelBtn.addEventListener("click", () => {
    cancelBtn.disabled = true;
    activePdfSession = activePdfSession ? { ...activePdfSession, cancelled: true } : null;
    persistOcrSession(activePdfSession);
    bubble.remove();
  });
}

async function askFileWithFallback({ prompt, file, thinkingBubble, signal }) {
  throwIfAborted(signal);
  if (isPdfFile(file)) {
    return runPdfOcrFlow({ file, prompt, pageCap: OCR_PAGE_LIMIT, thinkingBubble, signal });
  }
  const fd = new FormData();
  fd.append("message", prompt);
  fd.append("file", file);
  const response = await fetch("/api/ask-file", { method: "POST", body: fd, signal });
  const payload = await response.json().catch(() => ({}));
  throwIfAborted(signal);
  if (!response.ok) {
    const msg = payload?.error || payload?.extractionError || "Ask-file request failed.";
    throw new Error(msg);
  }
  const answer = payload?.answer || "(no answer returned)";
  const warning = payload?.extractionStatus === "empty" ? "Extraction returned minimal text." : "";
  return { answer, warning, continueRange: null };
}

function setLibrarySelection(doc) {
  activeLibraryDoc = doc;
  if (librarySelect) {
    if (doc?.docId) {
      const existing = Array.from(librarySelect.options || []).some(opt => opt.value === doc.docId);
      if (!existing) {
        const opt = document.createElement("option");
        opt.value = doc.docId;
        const existingLabels = new Set(Array.from(librarySelect.options || []).map(opt => opt.textContent || ""));
        const baseLabel = getLectureDisplayLabel(doc) || doc.title || doc.key || doc.docId;
        opt.textContent = ensureUniqueLabel(baseLabel, existingLabels);
        opt.dataset.dynamic = "1";
        librarySelect.appendChild(opt);
      }
      librarySelect.value = doc.docId;
    } else {
      librarySelect.value = "";
    }
  }
  if (metaLectureSelect) {
    if (doc?.docId) {
      const existing = Array.from(metaLectureSelect.options || []).some(opt => opt.value === doc.docId);
      if (!existing) {
        const opt = document.createElement("option");
        opt.value = doc.docId;
        const existingLabels = new Set(Array.from(metaLectureSelect.options || []).map(opt => opt.textContent || ""));
        const baseLabel = getLectureDisplayLabel(doc) || doc.title || doc.key || doc.docId;
        opt.textContent = ensureUniqueLabel(baseLabel, existingLabels);
        opt.dataset.dynamic = "1";
        metaLectureSelect.appendChild(opt);
      }
      metaLectureSelect.value = doc.docId;
    } else {
      metaLectureSelect.value = "";
    }
  }
  updateLibraryUiForSelection(doc);
  activeConversationMeta = {
    ...activeConversationMeta,
    selectedDocId: doc?.docId || null,
    selectedDocTitle: getLectureDisplayLabel(doc) || doc?.title || doc?.key || "",
  };
  schedulePersistConversation(200);
  updateResponseContextBar();
  refreshMetaDataPanel();
}

function renderLibraryResults(records) {
  librarySearchResults = Array.isArray(records) ? records : [];
  if (!libraryResults) return;
  libraryResults.innerHTML = "";
  if (!records?.length) {
    const empty = document.createElement("li");
    empty.className = "library-empty";
    empty.textContent = "No matching lectures yet.";
    libraryResults.appendChild(empty);
    return;
  }
  records.forEach((rec) => {
    const li = document.createElement("li");
    li.className = "library-result";
    const statusClass = rec.status === "ready" ? "ready" : rec.status === "needs_browser_ocr" ? "warn" : "missing";
    const label = getLectureDisplayLabel(rec) || rec.title || rec.key || rec.docId;
    li.innerHTML = `
      <div class="library-row-top">
        <h4>${escapeHTML(label)}</h4>
        <span class="status-tag ${statusClass}">${rec.status || "missing"}</span>
      </div>
      <div class="library-meta">${escapeHTML(rec.bucket || "")} · ${escapeHTML(rec.key || "")}</div>
    `;
    const actions = document.createElement("div");
    actions.className = "library-actions-row";
    if (rec.status === "ready") {
      const useBtn = document.createElement("button");
      useBtn.className = "chip-btn";
      useBtn.type = "button";
      useBtn.textContent = "Use cached lecture";
      useBtn.addEventListener("click", async () => {
        setLibrarySelection({ ...rec, status: "ready" });
        setStatus(libraryStatus, "success", `Using cached lecture: ${label}`);
        if (!chatInput.value.trim() || chatStreaming) return;
        setChatBusy(true);
        try {
          await askLibraryDocQuestion({ ...rec, status: "ready" }, chatInput.value.trim());
        } catch {
          // handled upstream
        } finally {
          setChatBusy(false);
        }
      });
      actions.appendChild(useBtn);
    }
    const ingestBtn = document.createElement("button");
    ingestBtn.className = "chip-btn";
    ingestBtn.type = "button";
    ingestBtn.textContent = rec.status === "ready" ? "Reprocess" : rec.status === "needs_browser_ocr" ? "Finish OCR once" : "Process once";
    ingestBtn.addEventListener("click", () => triggerLibraryIngest(rec));
    actions.appendChild(ingestBtn);
    li.appendChild(actions);
    libraryResults.appendChild(li);
  });
}

async function runLibrarySearch() {
  const query = librarySearchInput?.value?.trim() || "";
  const params = new URLSearchParams();
  if (query) params.set("q", query);
  params.set("limit", "12");
  setStatus(libraryStatus, "pending", query ? `Searching "${query}"...` : "Loading library...");
  try {
    const resp = await fetch(`/api/library/search?${params.toString()}`);
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      const msg = payload?.error || payload?.details || "Library search failed.";
      throw new Error(msg);
    }
    const results = Array.isArray(payload?.results) ? payload.results : [];
    renderLibraryResults(results);
    if (results.length) {
      setStatus(libraryStatus, "success", `Found ${results.length} match${results.length === 1 ? "" : "es"}.`);
    } else {
      setStatus(libraryStatus, "info", query ? "No matches yet." : "No indexed lectures yet.");
    }
    return results;
  } catch (error) {
    console.error(error);
    renderLibraryResults([]);
    setStatus(libraryStatus, "error", error instanceof Error ? error.message : "Library search failed.");
    throw error;
  }
}

async function askLibraryDocQuestion(doc, prompt, controllerOverride = null) {
  const question = prompt.trim();
  if (!question) {
    chatInput.focus();
    setStatus(libraryStatus, "info", "Type a question for this lecture.");
    return;
  }
  if (doc?.status && doc.status !== "ready") {
    setStatus(libraryStatus, "error", "OCR still processing… please wait or refresh index.");
    appendChatMessage("assistant", "OCR is still processing. Please retry once it is ready.", { track: false }).classList.add("error");
    return;
  }
  if (DEBUG_STREAMING) {
    console.log("[LIBRARY][ASK]", { docId: doc.docId, key: doc.key, requestId });
  }
  const requestId = `lib-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const isTable = isTableIntent(question);
  const allowExcerpts = allowExcerptsFromPrompt(question);
  const cleanedMeta = cleanPromptToSentence(question);
  const userBubble = appendChatMessage("user", question, {
    rawPrompt: question,
    cleanedPrompt: cleanedMeta.cleaned,
    topics: cleanedMeta.topics,
  });
  renderMathIfReady(userBubble);
  const controller = controllerOverride || new AbortController();
  const { signal } = controller;
  const thinking = showThinkingBubble();
  registerOwenAbortController(controller, thinking);
  setBubbleStatus(thinking, "O.W.E.N. Is Thinking");
  let machineTxtUnavailable = false;
  try {
    const useMachineTxt = doc?.sourceType === "machine_txt" || (typeof doc?.txtKey === "string" && doc.txtKey.startsWith("machine/txt/"));
    const requestPayload = { docId: doc.docId, question, requestId };
    if (useMachineTxt) {
      if (typeof doc?.txtKey === "string" && doc.txtKey.trim()) {
        requestPayload.txtKey = doc.txtKey.trim();
      }
      requestPayload.sourceType = "machine_txt";
      const lectureTitle = typeof doc?.displayName === "string" ? doc.displayName.trim() : "";
      if (lectureTitle) {
        requestPayload.lectureTitle = lectureTitle;
      }
    }
    const resp = await fetch("/api/library/ask", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(requestPayload),
      signal,
    });
    const payload = await resp.json().catch(() => ({}));
    throwIfAborted(signal);
    const normalized = normalizeLibraryAskPayload(payload, allowExcerpts);
    const machineTxtFailure = useMachineTxt && isMachineTxtLoadFailure(payload);
    if (DEBUG_STREAMING) {
      console.log("[LIBRARY][PAYLOAD_FINAL]", {
        status: resp.status,
        keys: Object.keys(payload || {}),
        ok: normalized.ok,
        versioned: normalized.versioned,
        answerLen: (normalized.answer || "").length,
        evidenceCount: normalized.evidence?.length || 0,
        references: normalized.references?.length || 0,
        requestId: normalized.debug?.requestId,
      });
    }
    if (!resp.ok) {
      const msg =
        (machineTxtFailure ? "Unable to load lecture TXT" : "") ||
        (normalized.error && normalized.error.message) ||
        payload?.error ||
        payload?.details ||
        `Library Q&A failed (status ${resp.status}).`;
      if (machineTxtFailure) {
        machineTxtUnavailable = true;
        libraryPrimedDocId = null;
        updateLibraryUiForSelection(doc);
      }
      throw new Error(msg);
    }
    if (thinking && thinking.isConnected) thinking.remove();
    if (!normalized.ok) {
      const reqId = normalized.debug?.requestId || payload?.requestId;
      const fallback = (normalized.answer && normalized.answer.trim()) || (payload?.answer && String(payload.answer).trim()) || "No answer returned (check server logs)";
      const bubble = appendChatMessage("assistant", reqId ? `${fallback} (req ${reqId})` : fallback, { track: false });
      bubble.classList.add("error");
      setStatus(libraryStatus, "error", "Lecture answer was empty.");
      return payload;
    }
    const trimmedAnswer = (normalized.answer || "").trim();
    if (!trimmedAnswer) {
      const friendly = payload?.error?.message || payload?.error || "Library returned an empty answer.";
      const bubble = appendChatMessage("assistant", friendly, { track: false });
      bubble.classList.add("error");
      setStatus(libraryStatus, "error", friendly);
      return payload;
    }
    const answer = sanitizeAnswerText(trimmedAnswer || "(no answer returned)", { allowExcerpts }) || trimmedAnswer;
    const assistant = appendChatMessage("assistant", "", { model: "library", track: false });
    const msgId = assistant.dataset.msgId || generateMessageId();
    assistant.dataset.msgId = msgId;
    assistant.dataset.docId = doc.docId;
    assistant.dataset.requestId = normalized.debug?.requestId || payload?.requestId || requestId;
    messageState.set(msgId, { streamText: normalized.answer, lastAnswer: normalized.answer });
    if (DEBUG_STREAMING) {
      console.log("[LIBRARY][FINALIZE_PRE]", {
        msgId,
        docId: doc.docId,
        answerLen: answer.length,
        evidenceCount: normalized.evidence?.length || 0,
        requestId: normalized.debug?.requestId,
      });
    }
    if (DEBUG_LECTURE_UI) {
      console.log("[LIBRARY_UI][STORE]", {
        msgId,
        docId: doc.docId,
        payloadKeys: Object.keys(payload || {}),
        answerLen: (normalized.answer || "").length,
      });
      if ((normalized.answer || "").trim().length > 0 && (!answer || !answer.trim())) {
        console.error("[INV-2_FAIL]", {
          payloadAnswerLen: (normalized.answer || "").length,
          msgTextLen: (answer || "").length,
          requestId: normalized.debug?.requestId || payload?.requestId,
          docId: doc.docId,
        });
      }
      if ((normalized.answer || "").trim().length === 0) {
        console.warn("[LIBRARY_UI][WARN_EMPTY_ANSWER]", {
          msgId,
          docId: doc.docId,
          requestId: normalized.debug?.requestId || payload?.requestId,
        });
      }
    }
    renderWithTypewriter(assistant, answer, new Map(), {
      references: !isTable ? normalized.references : undefined,
      evidence: !isTable && allowExcerpts ? normalized.evidence : undefined,
      debug: normalized.debug,
    });
    logMessage("assistant", answer, {
      model: "library",
      id: msgId,
      docId: doc.docId,
      docTitle: getLectureDisplayLabel(doc) || doc.title || doc.key || doc.docId,
      requestId: normalized.debug?.requestId || payload?.requestId || requestId,
      references: !isTable ? normalized.references : undefined,
      evidence: !isTable && allowExcerpts ? normalized.evidence : undefined,
    });
    const label = getLectureDisplayLabel(doc) || doc.title || doc.docId;
    setStatus(libraryStatus, "success", `Answered using ${label}.`);
    setLibrarySelection({ ...doc, status: "ready" });
    chatInput.value = "";
    return payload;
  } catch (error) {
    if (isAbortError(error)) {
      if (thinking && thinking.isConnected) clearThinkingBubble(thinking);
      return;
    }
    console.error(error);
    if (thinking && thinking.isConnected) thinking.remove();
    if (machineTxtUnavailable) {
      libraryPrimedDocId = null;
      updateLibraryUiForSelection(doc);
    }
    const bubble = appendChatMessage("assistant", error instanceof Error ? error.message : "Unable to answer from library.", { track: false });
    bubble.classList.add("error");
    setStatus(libraryStatus, "error", error instanceof Error ? error.message : "Library Q&A failed.");
    throw error;
  } finally {
    clearOwenAbortController(controller);
    void recordLectureAnalytics(doc);
  }
}

function parseReferenceMarkers(text) {
  const cleaned = (text || "").replace(/\r/g, "\n").trim();
  if (!cleaned) return [];
  const matches = Array.from(cleaned.matchAll(/\b(slide|page)\s*(\d{1,4})\s*(?:[-–:]\s*)?/gi));
  if (!matches.length) return [];
  return matches
    .map((match, index) => {
      const number = Number(match[2]);
      if (!Number.isFinite(number)) return null;
      const start = (match.index ?? 0) + match[0].length;
      const end = index + 1 < matches.length ? (matches[index + 1].index ?? cleaned.length) : cleaned.length;
      let detail = cleaned.slice(start, end).replace(/\s+/g, " ").trim();
      detail = detail.replace(/^[-–:\u2022\s]+/, "").trim();
      return { type: match[1].toLowerCase(), number, text: detail };
    })
    .filter(Boolean);
}

function clipReferenceText(text, limit = 220) {
  const trimmed = (text || "").trim();
  if (!trimmed) return "";
  if (trimmed.length <= limit) return trimmed;
  return `${trimmed.slice(0, Math.max(0, limit - 3)).trim()}...`;
}

function normalizeReferenceEntries(references, { inlineReferences = [], docId } = {}) {
  const items = [];
  const fallbackLines = [];

  const pushItem = (item) => {
    if (!item) return;
    const existing = items.find(entry => entry.type === item.type && entry.number === item.number);
    if (existing) {
      if (!existing.text && item.text) existing.text = item.text;
      return;
    }
    items.push(item);
  };

  const parseText = (text) => {
    if (!text) return false;
    const parsed = parseReferenceMarkers(text);
    if (!parsed.length) return false;
    parsed.forEach(pushItem);
    return true;
  };

  const handleEntry = (entry) => {
    if (!entry) return;
    if (typeof entry === "string") {
      if (!parseText(entry)) fallbackLines.push(entry.trim());
      return;
    }
    if (typeof entry !== "object") return;
    if (entry.docId && docId && entry.docId !== docId) return;
    const slideNum = typeof entry.slide === "number" ? entry.slide : Number(entry.slide);
    const pageNum = typeof entry.page === "number" ? entry.page : Number(entry.page);
    const slide = Number.isFinite(slideNum) ? slideNum : null;
    const page = Number.isFinite(pageNum) ? pageNum : null;
    const type = slide ? "slide" : page ? "page" : null;
    const number = slide || page;
    const text = typeof entry.title === "string" ? entry.title.trim() : typeof entry.text === "string" ? entry.text.trim() : "";
    if (type && Number.isFinite(number)) {
      pushItem({ type, number, text });
      return;
    }
    const label = typeof entry.label === "string" ? entry.label.trim() : "";
    if (!parseText(label || text)) {
      const fallback = label || text;
      if (fallback) fallbackLines.push(fallback);
    }
  };

  if (Array.isArray(references) && references.length) {
    references.forEach(handleEntry);
  }

  if (!items.length && Array.isArray(inlineReferences) && inlineReferences.length) {
    inlineReferences.forEach(ref => {
      if (!parseText(ref)) fallbackLines.push((ref || "").trim());
    });
  }

  const fallbackText = fallbackLines.filter(Boolean).join("\n");
  return { items, fallbackText };
}

function buildReferencePanel({ items, fallbackText, label }) {
  const panelId = `ref-panel-${Math.random().toString(36).slice(2, 8)}`;
  const panel = document.createElement("div");
  panel.className = "reference-panel";
  const header = document.createElement("div");
  header.className = "reference-panel__header";
  const title = document.createElement("div");
  title.className = "reference-panel__title";
  title.textContent = label;
  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "reference-panel__toggle";
  toggle.textContent = "Show";
  toggle.setAttribute("aria-expanded", "false");

  const body = document.createElement("div");
  body.className = "reference-panel__body";
  body.hidden = true;
  body.id = `${panelId}-body`;
  toggle.setAttribute("aria-controls", body.id);
  header.append(title, toggle);

  let detail = null;
  let detailTitle = null;
  let detailText = null;
  let chips = [];
  let activeIndex = null;

  if (items.length) {
    const chipWrap = document.createElement("div");
    chipWrap.className = "reference-panel__chips";
    const detailWrap = document.createElement("div");
    detailWrap.className = "reference-panel__detail";
    detailWrap.hidden = true;
    detailWrap.id = `${panelId}-detail`;
    detailWrap.setAttribute("role", "region");
    detailWrap.setAttribute("aria-live", "polite");
    detailTitle = document.createElement("div");
    detailTitle.className = "reference-panel__detail-title";
    detailText = document.createElement("div");
    detailText.className = "reference-panel__detail-text";
    detailWrap.append(detailTitle, detailText);
    detail = detailWrap;

    chips = items.map((item, index) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "reference-chip";
      const prefix = item.type === "page" ? "Page" : "Slide";
      const labelText = Number.isFinite(item.number) ? `${prefix} ${item.number}` : prefix;
      chip.textContent = labelText;
      chip.setAttribute("aria-expanded", "false");
      chip.setAttribute("aria-controls", detailWrap.id);
      if (item.text) {
        chip.title = `${labelText} – ${clipReferenceText(item.text, 160)}`;
      }
      chip.addEventListener("click", () => {
        const isActive = activeIndex === index;
        activeIndex = isActive ? null : index;
        chips.forEach((btn, idx) => {
          const nextActive = idx === activeIndex;
          btn.classList.toggle("is-active", nextActive);
          btn.setAttribute("aria-expanded", String(nextActive));
        });
        if (!detail || !detailTitle || !detailText) return;
        if (activeIndex == null) {
          detail.hidden = true;
          detailTitle.textContent = "";
          detailText.textContent = "";
          return;
        }
        const active = items[activeIndex];
        const activeLabel = Number.isFinite(active.number) ? `${active.type === "page" ? "Page" : "Slide"} ${active.number}` : "Reference";
        detailTitle.textContent = activeLabel;
        detailText.textContent = active.text ? clipReferenceText(active.text, 260) : "Details unavailable.";
        detail.hidden = false;
      });
      return chip;
    });
    chipWrap.append(...chips);
    body.append(chipWrap, detailWrap);
  } else if (fallbackText) {
    const fallback = document.createElement("div");
    fallback.className = "reference-panel__fallback";
    fallback.textContent = fallbackText;
    body.appendChild(fallback);
  }

  const setExpanded = (expanded) => {
    body.hidden = !expanded;
    toggle.textContent = expanded ? "Hide" : "Show";
    toggle.setAttribute("aria-expanded", String(expanded));
    if (!expanded && detail) {
      detail.hidden = true;
      activeIndex = null;
      chips.forEach(btn => {
        btn.classList.remove("is-active");
        btn.setAttribute("aria-expanded", "false");
      });
    }
  };

  toggle.addEventListener("click", () => {
    setExpanded(body.hidden);
  });

  setExpanded(false);
  panel.append(header, body);
  return panel;
}

function renderReferenceChips(bubble, references) {
  if (!bubble) return;
  const existing = bubble.querySelector(".reference-panel");
  if (existing) existing.remove();
  const docId = bubble?.dataset?.docId || "";
  const responseCard = bubble.querySelector(".response-card");
  const inlineReferences = responseCard ? getResponseState(responseCard)?.inlineReferences?.items : [];
  const { items, fallbackText } = normalizeReferenceEntries(references, { inlineReferences, docId });
  if (!items.length && !fallbackText) return;
  const hasExplicitRefs = Array.isArray(references) && references.length > 0;
  if (!docId && !hasExplicitRefs) return;
  const label = items.length && items.every(item => item.type === "slide") ? "Slides" : "References";
  const panel = buildReferencePanel({
    items,
    fallbackText,
    label: items.length ? `${label} (${items.length})` : "References",
  });
  bubble.appendChild(panel);
}

function renderEvidenceToggle(bubble, evidence) {
  if (!bubble || !Array.isArray(evidence) || !evidence.length) return;
  const wrapper = document.createElement("div");
  wrapper.className = "evidence-wrapper";
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "chip-btn";
  btn.textContent = "Show evidence";
  const list = document.createElement("div");
  list.className = "evidence-list";
  list.hidden = true;
  evidence.forEach(entry => {
    if (!entry) return;
    const msgDocId = bubble?.dataset?.docId;
    if (entry.docId && msgDocId && entry.docId !== msgDocId) return;
    const row = document.createElement("div");
    row.className = "evidence-row";
    const ref = document.createElement("div");
    ref.className = "evidence-ref";
    const hasSlide = typeof entry.slide === "number";
    const hasPage = typeof entry.page === "number";
    const prefix = hasPage ? "Page" : "Slide";
    const number = hasPage ? entry.page : entry.slide;
    const label = number ? `${prefix} ${number}` : prefix;
    ref.textContent = entry.title ? `${label} – ${entry.title}` : label;
    const text = document.createElement("div");
    text.className = "evidence-text";
    text.textContent = (entry.excerpt || entry.text || "").slice(0, 800);
    row.appendChild(ref);
    row.appendChild(text);
    list.appendChild(row);
  });
  btn.addEventListener("click", () => {
    const nextHidden = !list.hidden;
    list.hidden = nextHidden;
    btn.textContent = nextHidden ? "Show evidence" : "Hide evidence";
  });
  wrapper.appendChild(btn);
  wrapper.appendChild(list);
  bubble.appendChild(wrapper);
}

async function runLibraryBrowserOcr({ docId, bucket, key, downloadUrl, title, pageCount }) {
  const label = title || key || docId;
  setStatus(libraryStatus, "pending", "Downloading PDF for OCR…");
  const url = downloadUrl || `/api/file?bucket=${encodeURIComponent(bucket)}&key=${encodeURIComponent(key)}`;
  const res = await fetch(url);
  if (!res.ok) {
    const msg = await res.text().catch(() => "");
    throw new Error(msg || "Could not download PDF for OCR.");
  }
  const blob = await res.blob();
  const arrayBuffer = await blob.arrayBuffer();
  const pdfjsLib = await loadPdfJsLib();
  const pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const totalPages = pageCount || pdfDoc?.numPages || 0;
  if (!totalPages) {
    throw new Error("PDF contains no pages to OCR.");
  }
  const maxPages = fullDocumentMode ? totalPages : Math.min(OCR_PAGE_LIMIT, totalPages);
  const pageNumbers = Array.from({ length: maxPages }, (_, i) => i + 1);
  const pending = [];
  let processed = 0;
  let lastFinalize = null;

  await promisePool(pageNumbers, OCR_CONCURRENCY, async (pageNumber) => {
    const render = await renderPdfPageToBlob(pdfjsLib, pdfDoc, pageNumber);
    const pageResult = await ocrPageImage({
      fileHash: docId,
      pageIndex: pageNumber - 1,
      blob: render.blob,
      filename: `${label}-p${pageNumber}.jpg`,
    });
    pending.push(pageResult);
    processed += 1;
    if (pending.length >= OCR_SAVE_BATCH) {
      lastFinalize = await finalizeOcrText({
        fileHash: docId,
        pages: [...pending],
        filename: label,
        totalPages,
      });
      pending.length = 0;
    }
    setStatus(libraryStatus, "pending", `OCR ${processed}/${pageNumbers.length} pages…`);
  });

  if (pending.length) {
    lastFinalize = await finalizeOcrText({
      fileHash: docId,
      pages: pending,
      filename: label,
      totalPages,
    });
  }

  setStatus(libraryStatus, "success", `OCR cached for ${label}.`);
  const readyDoc = { docId, bucket, key, title: label, status: "ready", extractedKey: lastFinalize?.extractedKey };
  libraryPrimedDocId = readyDoc.docId;
  setLibrarySelection(readyDoc);
  await runLibrarySearch();
  await loadLibraryList();
  return readyDoc;
}

async function triggerLibraryIngest(record) {
  if (!record) return;
  setStatus(libraryStatus, "pending", "Processing in Worker…");
  try {
    const resp = await fetch("/api/library/ingest", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ bucket: record.bucket, key: record.key }),
    });
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      const msg = payload?.message || payload?.error || payload?.details || "Library ingest failed.";
      const detail = payload?.stage ? ` (stage: ${payload.stage})` : "";
      throw new Error(`${msg}${detail}`);
    }
    if (payload.status === "needs_browser_ocr") {
      const readyDoc = await runLibraryBrowserOcr({
        docId: payload.docId || record.docId,
        bucket: payload.bucket || record.bucket,
        key: payload.key || record.key,
        downloadUrl: payload.downloadUrl,
        title: record.title,
        pageCount: payload.pageCount,
      });
      if (readyDoc?.docId) libraryPrimedDocId = readyDoc.docId;
      await runLibrarySearch();
      await loadLibraryList();
      return readyDoc;
    }
    const readyDoc = {
      docId: payload.docId || record.docId,
      title: record.title,
      bucket: record.bucket,
      key: record.key,
      status: "ready",
      extractedKey: payload.extractedKey || record.extractedKey,
    };
    libraryPrimedDocId = readyDoc.docId;
    setLibrarySelection(readyDoc);
    setStatus(libraryStatus, "success", "Cached text ready.");
    await runLibrarySearch();
    await loadLibraryList();
    if (chatInput.value.trim() && !chatStreaming) {
      setChatBusy(true);
      try {
        await askLibraryDocQuestion(readyDoc, chatInput.value.trim());
      } finally {
        setChatBusy(false);
      }
    }
    return readyDoc;
  } catch (error) {
    console.error(error);
    setStatus(libraryStatus, "error", error instanceof Error ? error.message : "Library ingest failed.");
    throw error;
  }
}

async function runLibraryBatchIndex() {
  setStatus(libraryStatus, "pending", "Rebuilding library index…");
  try {
    const resp = await fetch("/api/library/batch-index", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({}) });
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      const msg = payload?.error || payload?.details || "Batch index failed.";
      throw new Error(msg);
    }
    setStatus(libraryStatus, "success", `Indexed ${payload?.indexed ?? 0} PDFs.`);
    if (activeRoute === "faculty-pipeline") setPipelineLastRun(2);
    await runLibrarySearch();
    await loadLibraryList();
  } catch (error) {
    console.error(error);
    setStatus(libraryStatus, "error", error instanceof Error ? error.message : "Batch index failed.");
  }
}

async function runLibraryBatchIngest(mode = "embedded_only") {
  setStatus(libraryStatus, "pending", mode === "full" ? "Batch ingest (full)..." : "Batch ingest (embedded only)...");
  try {
    const resp = await fetch("/api/library/batch-ingest", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ mode }),
    });
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      const msg = payload?.error || payload?.details || "Batch ingest failed.";
      throw new Error(msg);
    }
    const queued = payload?.queued ? `Queued ${payload.queued} scanned PDFs.` : "";
    setStatus(
      libraryStatus,
      "success",
      `Processed ${payload?.processed ?? 0} docs. Ready: ${payload?.ready ?? 0}. Skipped: ${payload?.skipped ?? 0}. ${queued}`.trim(),
    );
    if (activeRoute === "faculty-pipeline") setPipelineLastRun(2);
    await runLibrarySearch();
    await loadLibraryList();
  } catch (error) {
    console.error(error);
    setStatus(libraryStatus, "error", error instanceof Error ? error.message : "Batch ingest failed.");
  }
}

async function loadLibraryList() {
  const hasTargets = Boolean(librarySelect || machineLectureSelect || metaLectureSelect || libraryManagerList);
  if (!hasTargets) return;
  if (librarySelect) {
    librarySelect.innerHTML = `<option value="">Loading lectures...</option>`;
  }
  updateMetaLectureOptions("loading");
  updateMachineLectureOptions("loading");
  if (libraryManagerStatus) {
    setStatus(libraryManagerStatus, "pending", "Loading lecture library...");
  }
  if (libraryManagerList) {
    libraryManagerList.innerHTML = `<div class="library-manager__empty">Loading lectures...</div>`;
  }
  try {
    const params = new URLSearchParams();
    const useMachineTxt = activeRoute === "student";
    if (useMachineTxt) {
      params.set("source", "machine_txt");
    } else {
      params.set("prefix", "library/");
      if (libraryManagerList) params.set("detail", "1");
    }
    const resp = await fetch(`/api/library/list?${params.toString()}`);
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(payload?.error || "Failed to load library list.");
    libraryListItems = Array.isArray(payload?.items) ? payload.items : [];
    libraryListItems.sort((a, b) => {
      const aLabel = getLectureDisplayLabel(a) || a.title || a.key || a.docId || "";
      const bLabel = getLectureDisplayLabel(b) || b.title || b.key || b.docId || "";
      return aLabel.localeCompare(bLabel);
    });
    applyLectureDisplayLabels(libraryListItems);
    if (libraryPendingDeleteId && !libraryListItems.some(item => item.docId === libraryPendingDeleteId)) {
      libraryPendingDeleteId = null;
    }
    if (librarySelect) {
      librarySelect.innerHTML = `<option value="">Select a lecture...</option>` + libraryListItems.map(item => {
        const label = getLectureDisplayLabel(item) || item.title || item.key || item.docId;
        return `<option value="${escapeHTML(item.docId)}">${escapeHTML(label)}</option>`;
      }).join("");
    }
    updateMetaLectureOptions();
    updateMachineLectureOptions();
    renderLibraryManagerList();
  } catch (err) {
    console.error(err);
    if (librarySelect) {
      librarySelect.innerHTML = `<option value="">Failed to load</option>`;
    }
    updateMetaLectureOptions("error");
    updateMachineLectureOptions("error");
    if (libraryManagerStatus) {
      setStatus(libraryManagerStatus, "error", err instanceof Error ? err.message : "Library load failed.");
    }
    if (libraryManagerList) {
      libraryManagerList.innerHTML = `<div class="library-manager__empty">Unable to load lectures.</div>`;
    }
  }
}

function isLibraryDocPrimed(doc) {
  return Boolean(doc?.docId && libraryPrimedDocId === doc.docId);
}

function updateLibraryUiForSelection(doc) {
  const isPipeline = activeRoute === "faculty-pipeline";
  const isPrimed = isLibraryDocPrimed(doc);
  if (libraryActionBtn) {
    libraryActionBtn.disabled = !doc;
    if (isPipeline) {
      libraryActionBtn.classList.remove("library-primed");
      libraryActionBtn.textContent = "Run OCR / Parse";
    } else {
      libraryActionBtn.classList.toggle("library-primed", isPrimed);
      libraryActionBtn.textContent = isPrimed ? "Lecture Primed" : "Prime Cached Lecture";
    }
  }
  if (libraryQuizBtn) {
    const canQuiz = Boolean(doc) && isPrimed;
    libraryQuizBtn.disabled = !canQuiz;
    const helperMessage = !doc ? "Select a lecture first." : !isPrimed ? "Prime the lecture first." : "";
    if (libraryQuizHelper) {
      libraryQuizHelper.textContent = helperMessage;
      libraryQuizHelper.hidden = !helperMessage;
    }
    libraryQuizBtn.title = helperMessage || "Generate a quiz from the cached lecture.";
  }
}

function onLibrarySelectChange() {
  if (!librarySelect) return;
  const docId = librarySelect.value;
  const selected = libraryListItems.find(item => item.docId === docId) || null;
  setLibrarySelection(selected || null);
}

async function onLibraryAction() {
  if (!activeLibraryDoc) return;
  if (activeRoute !== "faculty-pipeline" && activeLibraryDoc.status === "ready") {
    const label = getLectureDisplayLabel(activeLibraryDoc) || activeLibraryDoc.title || activeLibraryDoc.key;
    setStatus(libraryStatus, "success", `Using cached lecture: ${label}`);
    if (activeLibraryDoc.docId) {
      libraryPrimedDocId = activeLibraryDoc.docId;
      updateLibraryUiForSelection(activeLibraryDoc);
    }
    return;
  }
  const readyDoc = await triggerLibraryIngest(activeLibraryDoc);
  if (readyDoc?.docId) updateLibraryUiForSelection(readyDoc);
  if (readyDoc?.docId && activeRoute === "faculty-pipeline") {
    setPipelineLastRun(3);
  }
}

const QUIZ_BATCH_SIZE = 5;
const QUIZ_CHOICE_IDS = new Set(["A", "B", "C", "D", "E"]);
const QUIZ_DIFFICULTIES = new Set(["easy", "medium", "hard"]);

function setQuizOpen(isOpen) {
  if (isOpen) {
    quizLastActiveElement = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  }
  if (quizOverlay) quizOverlay.hidden = !isOpen;
  if (quizModal) {
    quizModal.hidden = !isOpen;
    quizModal.setAttribute("aria-hidden", String(!isOpen));
  }
  document.body.classList.toggle("quiz-open", isOpen);
  if (!isOpen && quizLastActiveElement && quizLastActiveElement.isConnected) {
    quizLastActiveElement.focus();
  }
  if (!isOpen) {
    quizLastActiveElement = null;
  }
}

function setQuizError(message) {
  if (!quizError || !quizErrorText) return;
  if (!message) {
    quizError.hidden = true;
    quizErrorText.textContent = "";
    return;
  }
  quizError.hidden = false;
  quizErrorText.textContent = message;
}

function setQuizGenerating(next) {
  quizGenerating = next;
  if (quizLoading) quizLoading.hidden = !next;
  if (quizInterruptBtn) {
    quizInterruptBtn.hidden = !next;
    quizInterruptBtn.disabled = !next || quizInterrupting;
  }
  if (quizGenerateMoreBtn) quizGenerateMoreBtn.disabled = next;
  if (quizRetryBtn) quizRetryBtn.disabled = next;
}

function abortQuizGeneration() {
  const wasGenerating = quizGenerating || Boolean(quizAbortController);
  if (quizAbortController) {
    quizAbortController.abort();
    quizAbortController = null;
  }
  setQuizGenerating(false);
  return wasGenerating;
}

function buildQuizSession(doc) {
  const label = getLectureDisplayLabel(doc) || doc?.title || doc?.key || doc?.docId || "";
  return {
    token: `${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    lectureId: doc?.docId || "",
    lectureTitle: label,
    sourceType: doc?.sourceType || "",
    txtKey: doc?.txtKey || "",
    sets: [],
    totalPoints: 0,
    totalAnswered: 0,
  };
}

function resetQuizUI() {
  quizInterrupting = false;
  if (quizLectureTitle) quizLectureTitle.textContent = "";
  if (quizSets) quizSets.innerHTML = "";
  if (quizGenerateMoreBtn) quizGenerateMoreBtn.hidden = true;
  setQuizError("");
  setQuizGenerating(false);
}

function closeQuiz({ abort = quizGenerating || Boolean(quizAbortController) } = {}) {
  if (abort) abortQuizGeneration();
  quizSession = null;
  resetQuizUI();
  setQuizOpen(false);
}

async function interruptQuizSession() {
  if (quizInterrupting) return;
  if (!quizSession) {
    setQuizError("No active quiz to interrupt.");
    return;
  }
  quizInterrupting = true;
  if (quizInterruptBtn) quizInterruptBtn.disabled = true;
  setQuizError("");
  const sessionToken = quizSession.token;
  abortQuizGeneration();
  try {
    const resp = await fetch("/api/library/quiz/interrupt", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        lectureId: quizSession.lectureId,
        lectureTitle: quizSession.lectureTitle || "",
        token: sessionToken,
      }),
    });
    const { payload, rawText } = await readJsonResponse(resp);
    if (!resp.ok || payload?.ok === false) {
      const message = extractErrorMessage(payload, rawText, "Quiz interrupt failed.");
      throw new Error(message);
    }
    if (quizSession?.token !== sessionToken) return;
    closeQuiz({ abort: false });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Quiz interrupt failed.";
    setQuizError(message);
  } finally {
    quizInterrupting = false;
    if (quizInterruptBtn && quizGenerating) {
      quizInterruptBtn.disabled = false;
    }
  }
}

function openQuizForLecture(doc) {
  abortQuizGeneration();
  quizSession = buildQuizSession(doc);
  if (quizLectureTitle) {
    quizLectureTitle.textContent = quizSession.lectureTitle || "Selected lecture";
  }
  if (quizSets) quizSets.innerHTML = "";
  if (quizGenerateMoreBtn) quizGenerateMoreBtn.hidden = true;
  setQuizError("");
  setQuizGenerating(false);
  setQuizOpen(true);
}

function buildQuizRequestPayload(session) {
  const payload = {
    docId: session.lectureId,
    lectureTitle: session.lectureTitle || "",
  };
  if (session.sourceType) payload.sourceType = session.sourceType;
  if (session.txtKey) payload.txtKey = session.txtKey;
  return payload;
}

function normalizeQuizChoiceId(value) {
  return String(value || "").trim().toUpperCase();
}

function normalizeQuizDifficulty(value) {
  return String(value || "").trim().toLowerCase();
}

function validateQuizBatchShape(batch) {
  const errors = [];
  if (!batch || typeof batch !== "object") {
    return ["Quiz batch is not an object."];
  }
  const lectureTitle = typeof batch.lectureTitle === "string" ? batch.lectureTitle.trim() : "";
  if (!lectureTitle) errors.push("lectureTitle is required.");
  const setSize = typeof batch.setSize === "number" && Number.isFinite(batch.setSize) ? batch.setSize : NaN;
  if (!Number.isFinite(setSize)) {
    errors.push("setSize is required.");
  } else if (setSize !== QUIZ_BATCH_SIZE) {
    errors.push(`setSize must be ${QUIZ_BATCH_SIZE}.`);
  }
  const questions = batch.questions;
  if (!Array.isArray(questions)) {
    errors.push("questions must be an array.");
    return errors;
  }
  if (questions.length !== QUIZ_BATCH_SIZE) {
    errors.push(`questions must include exactly ${QUIZ_BATCH_SIZE} items.`);
  }
  if (Number.isFinite(setSize) && questions.length !== setSize) {
    errors.push("questions length must match setSize.");
  }
  const questionIds = new Set();
  questions.forEach((question, index) => {
    if (!question || typeof question !== "object") {
      errors.push(`questions[${index}] is not an object.`);
      return;
    }
    const id = typeof question.id === "string" ? question.id.trim() : "";
    if (!id) {
      errors.push(`questions[${index}].id is required.`);
    } else if (questionIds.has(id)) {
      errors.push(`questions[${index}].id is duplicated.`);
    } else {
      questionIds.add(id);
    }
    if (typeof question.stem !== "string" || !question.stem.trim()) {
      errors.push(`questions[${index}].stem is required.`);
    }
    if (typeof question.rationale !== "string" || !question.rationale.trim()) {
      errors.push(`questions[${index}].rationale is required.`);
    }
    const tags = question.tags;
    if (!Array.isArray(tags) || !tags.length) {
      errors.push(`questions[${index}].tags must be a non-empty array.`);
    } else {
      tags.forEach((tag, tagIndex) => {
        if (typeof tag !== "string" || !tag.trim()) {
          errors.push(`questions[${index}].tags[${tagIndex}] must be a string.`);
        }
      });
    }
    const difficulty = normalizeQuizDifficulty(question.difficulty);
    if (!difficulty || !QUIZ_DIFFICULTIES.has(difficulty)) {
      errors.push(`questions[${index}].difficulty must be easy, medium, or hard.`);
    }
    const choices = question.choices;
    if (!Array.isArray(choices)) {
      errors.push(`questions[${index}].choices must be an array.`);
    } else {
      if (choices.length < 4 || choices.length > 5) {
        errors.push(`questions[${index}].choices must have 4 or 5 items.`);
      }
      const choiceIds = new Set();
      choices.forEach((choice, choiceIndex) => {
        if (!choice || typeof choice !== "object") {
          errors.push(`questions[${index}].choices[${choiceIndex}] is not an object.`);
          return;
        }
        const choiceId = normalizeQuizChoiceId(choice.id);
        if (!choiceId || !QUIZ_CHOICE_IDS.has(choiceId)) {
          errors.push(`questions[${index}].choices[${choiceIndex}].id must be A-E.`);
        } else if (choiceIds.has(choiceId)) {
          errors.push(`questions[${index}].choices has duplicate id ${choiceId}.`);
        } else {
          choiceIds.add(choiceId);
        }
        if (typeof choice.text !== "string" || !choice.text.trim()) {
          errors.push(`questions[${index}].choices[${choiceIndex}].text is required.`);
        }
      });
      const answer = normalizeQuizChoiceId(question.answer);
      if (!answer || !QUIZ_CHOICE_IDS.has(answer)) {
        errors.push(`questions[${index}].answer must be A-E.`);
      } else if (!choiceIds.has(answer)) {
        errors.push(`questions[${index}].answer must match a choice id.`);
      }
    }
    if (Object.prototype.hasOwnProperty.call(question, "references") && question.references !== undefined) {
      if (!Array.isArray(question.references) || !question.references.every(ref => typeof ref === "string")) {
        errors.push(`questions[${index}].references must be an array of strings.`);
      }
    }
  });
  return errors;
}

function formatQuizForUI(batch) {
  const questions = batch.questions.map(question => {
    const choices = Array.isArray(question?.choices)
      ? question.choices.map(choice => ({
        id: normalizeQuizChoiceId(choice?.id),
        text: String(choice?.text || "").trim(),
      }))
      : [];
    return {
      id: String(question?.id || "").trim(),
      stem: String(question?.stem || "").trim(),
      choices,
      correctChoiceId: normalizeQuizChoiceId(question?.answer),
      rationale: String(question?.rationale || "").trim(),
      tags: Array.isArray(question?.tags)
        ? question.tags.map(tag => String(tag || "").trim()).filter(Boolean)
        : [],
      difficulty: normalizeQuizDifficulty(question?.difficulty),
      references: Array.isArray(question?.references)
        ? question.references.map(ref => String(ref || "").trim()).filter(Boolean)
        : [],
      attemptsCount: 0,
      guessedChoiceIds: new Set(),
      incorrectChoiceIds: new Set(),
      isCompleted: false,
      firstAttemptCorrect: false,
    };
  });
  return {
    lectureTitle: String(batch.lectureTitle || "").trim(),
    setSize: batch.setSize,
    questions,
    scoring: { firstAttemptOnly: true },
    isCompleted: false,
    pointsEarned: 0,
  };
}

function renderQuizSet(setIndex, set) {
  const setEl = document.createElement("section");
  setEl.className = "quiz-set";
  setEl.dataset.quizSetIndex = String(setIndex);

  const questionsHtml = set.questions.map((question, questionIndex) => {
    const choicesHtml = question.choices.map(choice => {
      const choiceId = escapeHTML(choice.id || "");
      const choiceText = escapeHTML(choice.text || "");
      return `
        <button class="quiz-option" type="button" data-quiz-set-index="${setIndex}" data-quiz-question-index="${questionIndex}" data-choice-id="${choiceId}">
          <span class="quiz-option-bubble" aria-hidden="true"></span>
          <span class="quiz-option-text"><strong>${choiceId}.</strong> ${choiceText}</span>
        </button>
      `;
    }).join("");
    const referencesText = question.references.length
      ? `References: ${question.references.map(ref => escapeHTML(ref)).join(", ")}`
      : "";
    return `
      <div class="quiz-question" data-quiz-set-index="${setIndex}" data-quiz-question-index="${questionIndex}">
        <div class="quiz-question-title">Question ${questionIndex + 1}</div>
        <div class="quiz-question-stem">${escapeHTML(question.stem || "")}</div>
        <div class="quiz-options">${choicesHtml}</div>
        <div class="quiz-explanation" data-role="quiz-explanation" hidden>${escapeHTML(question.rationale || "")}</div>
        <div class="quiz-references" data-role="quiz-references" hidden>${referencesText}</div>
      </div>
    `;
  }).join("");

  setEl.innerHTML = `
    <div class="quiz-set-header">
      <span>Set ${setIndex + 1}</span>
      <span>${set.questions.length} Questions</span>
    </div>
    ${questionsHtml}
    <div class="quiz-set-results" data-quiz-set-results hidden>
      <div data-role="quiz-set-score"></div>
      <div data-role="quiz-overall-score"></div>
    </div>
  `;
  return setEl;
}

function updateOverallScores() {
  if (!quizSession || !quizSets) return;
  const total = quizSession.totalAnswered;
  const points = quizSession.totalPoints;
  const percent = total ? Math.round((points / total) * 100) : 0;
  quizSets.querySelectorAll("[data-role=\"quiz-overall-score\"]").forEach((node) => {
    node.textContent = `Overall Score: ${points}/${total} (${percent}%)`;
  });
}

function renderQuizSetResults(setIndex) {
  if (!quizSession || !quizSets) return;
  const set = quizSession.sets[setIndex];
  if (!set) return;
  const setEl = quizSets.querySelector(`.quiz-set[data-quiz-set-index="${setIndex}"]`);
  if (!setEl) return;
  const results = setEl.querySelector("[data-quiz-set-results]");
  if (!results) return;
  const scoring = set.scoring || { firstAttemptOnly: true };
  const points = scoring.firstAttemptOnly
    ? set.questions.filter(question => question.firstAttemptCorrect).length
    : set.questions.filter(question => question.isCompleted).length;
  set.pointsEarned = points;
  const setPercent = set.questions.length ? Math.round((points / set.questions.length) * 100) : 0;
  const setScore = results.querySelector("[data-role=\"quiz-set-score\"]");
  if (setScore) {
    setScore.textContent = `Set Score: ${points}/${set.questions.length} (${setPercent}%)`;
  }
  updateOverallScores();
  results.hidden = false;
  if (quizGenerateMoreBtn) quizGenerateMoreBtn.hidden = false;
}

function updateQuizQuestionUI(setIndex, questionIndex) {
  if (!quizSession || !quizSets) return;
  const set = quizSession.sets[setIndex];
  const question = set?.questions[questionIndex];
  if (!question) return;
  const questionEl = quizSets.querySelector(`.quiz-question[data-quiz-set-index="${setIndex}"][data-quiz-question-index="${questionIndex}"]`);
  if (!questionEl) return;
  const options = Array.from(questionEl.querySelectorAll(".quiz-option"));
  options.forEach(option => {
    const choiceId = option.dataset.choiceId || "";
    if (question.incorrectChoiceIds.has(choiceId)) {
      option.classList.add("is-incorrect");
      option.disabled = true;
    }
    if (question.isCompleted) {
      option.disabled = true;
      if (choiceId === question.correctChoiceId) {
        option.classList.add("is-correct");
      }
    }
  });
  const explanation = questionEl.querySelector("[data-role=\"quiz-explanation\"]");
  if (explanation) explanation.hidden = !question.isCompleted;
  const refs = questionEl.querySelector("[data-role=\"quiz-references\"]");
  if (refs) {
    refs.hidden = !question.isCompleted || !question.references.length;
  }
}

function handleQuizOptionClick(event) {
  const target = event.target instanceof Element ? event.target.closest(".quiz-option") : null;
  if (!target || !quizSession) return;
  const setIndex = Number(target.dataset.quizSetIndex || -1);
  const questionIndex = Number(target.dataset.quizQuestionIndex || -1);
  const choiceId = target.dataset.choiceId || "";
  const set = quizSession.sets[setIndex];
  const question = set?.questions?.[questionIndex];
  if (!question || question.isCompleted) return;
  const scoring = set?.scoring || { firstAttemptOnly: true };
  const wasCompleted = question.isCompleted;
  question.attemptsCount += 1;
  question.guessedChoiceIds.add(choiceId);
  if (choiceId === question.correctChoiceId) {
    question.isCompleted = true;
    question.firstAttemptCorrect = question.attemptsCount === 1;
  } else {
    question.incorrectChoiceIds.add(choiceId);
  }
  updateQuizQuestionUI(setIndex, questionIndex);
  if (!wasCompleted && question.isCompleted) {
    quizSession.totalAnswered += 1;
    const earned = scoring.firstAttemptOnly ? question.firstAttemptCorrect : true;
    if (earned) quizSession.totalPoints += 1;
    const setCompleted = set.questions.every(q => q.isCompleted);
    if (setCompleted && !set.isCompleted) {
      set.isCompleted = true;
      renderQuizSetResults(setIndex);
    } else {
      updateOverallScores();
    }
  }
}

async function generateQuizBatch(session, controller) {
  const payload = buildQuizRequestPayload(session);
  const resp = await fetch("/api/library/quiz", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(payload),
    signal: controller.signal,
  });
  const { payload: data, rawText } = await readJsonResponse(resp);
  if (!resp.ok || data?.ok === false) {
    const msg = extractErrorMessage(data, rawText, "Quiz generation failed.");
    throw new Error(msg);
  }
  if (!data?.batch) {
    throw new Error("Quiz generation failed.");
  }
  return data.batch;
}

async function runQuizBatch() {
  if (quizGenerating || !quizSession) return;
  abortQuizGeneration();
  const sessionToken = quizSession.token;
  setQuizGenerating(true);
  setQuizError("");

  const controller = new AbortController();
  quizAbortController = controller;
  try {
    const rawBatch = await generateQuizBatch(quizSession, controller);
    const validationErrors = validateQuizBatchShape(rawBatch);
    if (validationErrors.length) {
      throw new Error("Quiz JSON failed validation.");
    }
    if (quizSession?.token !== sessionToken) return;
    const setIndex = quizSession.sets.length;
    const set = formatQuizForUI(rawBatch);
    quizSession.sets.push(set);
    if (quizSets) {
      quizSets.appendChild(renderQuizSet(setIndex, set));
    }
    if (quizGenerateMoreBtn) quizGenerateMoreBtn.hidden = true;
  } catch (error) {
    if (isAbortError(error)) return;
    if (quizSession?.token !== sessionToken) return;
    const message = error instanceof Error ? error.message : "Quiz generation failed.";
    setQuizError(message);
  } finally {
    if (quizAbortController === controller) {
      quizAbortController = null;
    }
    if (quizSession?.token !== sessionToken) return;
    setQuizGenerating(false);
  }
}

async function onLibraryQuizAction() {
  if (!activeLibraryDoc) {
    window.alert("Select a lecture first.");
    return;
  }
  if (!isLibraryDocPrimed(activeLibraryDoc)) {
    window.alert("Prime the lecture first.");
    return;
  }
  openQuizForLecture(activeLibraryDoc);
  await runQuizBatch();
}

function handleLocalCommands(prompt) {
  const match = prompt.match(/^\/file\s+(\S+)(?:\s+(\S+))?/i);
  if (match) {
    const key = match[1];
    const bucket = match[2] || "";
    retrieveFileFromR2({ key, bucket, autoDownload: true, statusEl: retrieveStatus, emitChat: true });
    return true;
  }
  return false;
}

const SUPPORTED_PASTE_TYPES = new Set(["application/pdf"]);

function fileExtFromMime(mime) {
  if (!mime) return "bin";
  if (mime.includes("pdf")) return "pdf";
  if (mime.includes("png")) return "png";
  if (mime.includes("jpeg")) return "jpeg";
  if (mime.includes("jpg")) return "jpg";
  if (mime.includes("webp")) return "webp";
  if (mime.includes("gif")) return "gif";
  return "bin";
}

function collectPastedFiles(event) {
  const files = [];
  const items = Array.from(event.clipboardData?.items || []);
  items.forEach(item => {
    if (item.kind !== "file") return;
    const file = item.getAsFile();
    if (!file) return;
    if (file.type.startsWith("image/") || SUPPORTED_PASTE_TYPES.has(file.type)) {
      files.push(file);
    }
  });
  if (!files.length && event.clipboardData?.files?.length) {
    Array.from(event.clipboardData.files).forEach(file => {
      if (file.type.startsWith("image/") || SUPPORTED_PASTE_TYPES.has(file.type)) {
        files.push(file);
      }
    });
  }
  return files;
}

function initBucketSelect() {
  if (!bucketSelect) return;
  bucketSelect.innerHTML = DESTINATIONS.map(dest => `<option value="${dest.value}">${dest.label}</option>`).join("");
  let saved = "";
  try {
    saved = localStorage.getItem(DEST_STORAGE_KEY) || "";
  } catch {
    saved = "";
  }
  const fallback = DESTINATIONS[0]?.value || "";
  bucketSelect.value = saved && DESTINATIONS.some(d => d.value === saved) ? saved : fallback;
  bucketSelect.addEventListener("change", () => {
    try {
      localStorage.setItem(DEST_STORAGE_KEY, bucketSelect.value);
    } catch {}
  });
}

function ensureNamedFile(file, index) {
  const hasName = file?.name && file.name !== "image.png";
  if (hasName) return file;
  const ext = fileExtFromMime(file?.type || "");
  const stamp = Date.now().toString(36);
  const name = `pasted-${stamp}-${index + 1}.${ext}`;
  return new File([file], name, { type: file?.type || "application/octet-stream" });
}

function handleChatPaste(event) {
  const files = collectPastedFiles(event);
  if (!files.length) return;
  files.forEach((file, idx) => {
    const normalized = ensureNamedFile(file, idx);
    attachments.length = 0;
    activePdfSession = null;
    clearOcrSession();
    if (oneShotPreviewUrl && oneShotPreviewUrl.startsWith("blob:")) URL.revokeObjectURL(oneShotPreviewUrl);
    oneShotFile = normalized;
    oneShotPreviewUrl = normalized.type.startsWith("image/") ? URL.createObjectURL(normalized) : "";
    updateAttachmentIndicator();
  });
}

async function runChat(prompt) {
  skipAllTypewriters();
  if (chatStreaming) return;
  if (handleLocalCommands(prompt)) return;
  setChatBusy(true);
  let live = null;
  let abortController = null;
  try {
    if (activeLibraryDoc && attachments.length === 0 && !oneShotFile) {
      try {
        await askLibraryDocQuestion(activeLibraryDoc, prompt);
      } catch {
        // Errors are already surfaced inside askLibraryDocQuestion.
      }
      return;
    }
    const model = currentModel;
    if (attachmentUploads.size) {
      attachStatus.dataset.state = "attached";
      attachStatus.innerHTML = `
        <span class="attachment-chip loading">
          <span>Finalizing attachments…</span>
          <span class="loading-track"></span>
        </span>`;
      await Promise.all(Array.from(attachmentUploads));
    }
    if (documentSummaryTasks.size) {
      attachStatus.dataset.state = "attached";
      attachStatus.innerHTML = `
        <span class="attachment-chip loading">
          <span>Analyzing document${documentSummaryTasks.size > 1 ? "s" : ""}…</span>
          <span class="loading-track"></span>
        </span>`;
      await Promise.all(Array.from(documentSummaryTasks));
    }
    const historyPayload = getHistoryPayload(10);
    const includeHistory = shouldAttachHistory(prompt, historyPayload);
    const promptTags = updateHashtagCloud(prompt);
    const promptMeta = cleanPromptToSentence(prompt);
    const excerptsAllowed = allowExcerptsFromPrompt(prompt);
    const attachmentsSnapshot = cloneMessageAttachments(attachments);
    const fileRefs = buildFileReferencesFromAttachments(attachmentsSnapshot);
    const messages = includeHistory
      ? historyPayload.map(entry => ({ role: entry.role, content: entry.text }))
      : [];
    messages.push({ role: "user", content: prompt });

    // One-shot path: file + question in same request, no R2 persistence.
    if (oneShotFile) {
      if (!isPdfFile(oneShotFile)) {
        activePdfSession = null;
      }
      const attachmentsForBubble = [
        {
          filename: oneShotFile.name,
          mimeType: oneShotFile.type,
          previewUrl: oneShotPreviewUrl || "",
        },
      ];
      appendChatMessage("user", prompt, {
        attachments: attachmentsForBubble,
        rawPrompt: prompt,
        cleanedPrompt: promptMeta.cleaned,
        topics: promptMeta.topics,
      });
      chatInput.value = "";
      if (insightsLog) insightsLog.textContent = `Last ask (${model}): ${prompt}`;
      const thinkingBubble = showThinkingBubble();
      abortController = registerOwenAbortController(new AbortController(), thinkingBubble);
      const { signal } = abortController;
      try {
        const result = await askFileWithFallback({ prompt, file: oneShotFile, thinkingBubble, signal });
        throwIfAborted(signal);
        if (thinkingBubble && thinkingBubble.isConnected) thinkingBubble.remove();
        const finalAnswer = result.warning ? `${result.answer}\n\n_${result.warning}_` : result.answer;
        const docId = result.fileHash || activePdfSession?.fileHash || null;
        const extractedKey = result.extractedKey || activePdfSession?.extractedKey || null;
        const answerBubble = appendChatMessage("assistant", "", { model: "gpt-5-mini", track: false, docId, extractedKey });
        renderWithTypewriter(answerBubble, finalAnswer, new Map());
        const msgId = answerBubble.dataset.msgId || generateMessageId();
        logMessage("assistant", finalAnswer, { model: "gpt-5-mini", id: msgId, docId, extractedKey });
        if (result?.continueRange) {
          renderContinuePrompt(result.continueRange, prompt);
        }
      } catch (error) {
        if (isAbortError(error)) {
          if (thinkingBubble && thinkingBubble.isConnected) clearThinkingBubble(thinkingBubble);
          return;
        }
        console.error(error);
        if (thinkingBubble && thinkingBubble.isConnected) thinkingBubble.remove();
        const bubble = appendChatMessage("assistant", error instanceof Error ? error.message : "Something went wrong.");
        bubble.classList.add("error");
      } finally {
        if (oneShotPreviewUrl && oneShotPreviewUrl.startsWith("blob:")) URL.revokeObjectURL(oneShotPreviewUrl);
        oneShotFile = null;
        oneShotPreviewUrl = "";
        updateAttachmentIndicator();
      }
      return;
    }

    const payload = {
      model,
      conversation_id: activeConversationId,
      messages,
    };
    if (Array.isArray(promptTags) && promptTags.length) {
      payload.meta_tags = promptTags;
    }
    if (fileRefs.length) {
      payload.files = fileRefs;
      payload.attachments = fileRefs;
      payload.fileRefs = fileRefs;
    }
    appendChatMessage("user", prompt, {
      attachments: attachmentsSnapshot,
      rawPrompt: prompt,
      cleanedPrompt: promptMeta.cleaned,
      topics: promptMeta.topics,
    });
    chatInput.value = "";
    if (insightsLog) insightsLog.textContent = `Last ask (${model}): ${prompt}`;
    const thinkingBubble = showThinkingBubble();
    abortController = registerOwenAbortController(new AbortController(), thinkingBubble);
    const { signal } = abortController;
    setBubbleStatus(thinkingBubble, "Queued…");
    live = thinkingBubble;
    let hasContent = false;
    let acc = "";
    const citationEvents = [];
    const msgId = thinkingBubble?.dataset?.msgId || generateMessageId();
    if (thinkingBubble) thinkingBubble.dataset.msgId = msgId;
    messageState.set(msgId, { streamText: "", lastAnswer: "" });

    setBubbleStatus(live, "O.W.E.N. Is Thinking");
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
      signal,
    });
    throwIfAborted(signal);
    if (isImageModel(model)) {
      const imagePayload = await response.json().catch(() => ({}));
      throwIfAborted(signal);
      if (!response.ok || imagePayload?.ok === false) {
        const errMessage = imagePayload?.error?.message || imagePayload?.error || imagePayload?.message;
        throw new Error(errMessage || "Image generation failed.");
      }
      const imageUrl = imagePayload.image_url || (imagePayload.image_base64 ? `data:image/png;base64,${imagePayload.image_base64}` : "");
      if (!imageUrl) {
        throw new Error("Sora did not return an image.");
      }
      const caption = imagePayload.revised_prompt || imagePayload.prompt || "Generated image";
      if (live && live.isConnected) live.remove();
      appendChatMessage("assistant", caption, {
        model,
        imageUrl,
        imageAlt: caption,
      });
      return;
    }
    const contentType = response.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
      const payload = await response.json().catch(() => ({}));
      throwIfAborted(signal);
      if (!response.ok || payload?.ok === false) {
        const errMessage = payload?.error?.message || payload?.error || payload?.message;
        throw new Error(errMessage || "Chat request failed.");
      }
      if (Array.isArray(payload?.answerSegments)) {
        if (live) {
          live.classList.remove("thinking", "scaffold");
        }
        const segments = Array.isArray(payload.answerSegments) ? payload.answerSegments : [];
        const sources = Array.isArray(payload.sources) ? payload.sources : [];
        const citationMap = buildCitationMapFromSegments(segments);
        const normalizedSources = normalizeCitedSources(sources, citationMap);
        const plainAnswer = buildCitedPlainText(segments);
        renderCitedAnswerSegments(live, segments, normalizedSources);
        messageState.set(msgId, { streamText: plainAnswer, lastAnswer: plainAnswer });
        logMessage("assistant", plainAnswer, {
          model,
          id: msgId,
          requestId: payload?.requestId,
          sources: normalizedSources,
          citations: normalizedSources,
          answerSegments: segments,
          renderedMarkdown: plainAnswer,
        });
        return;
      }
      const rawAnswer =
        typeof payload?.answer === "string"
          ? payload.answer
          : typeof payload?.output_text === "string"
            ? payload.output_text
            : typeof payload?.response === "string"
              ? payload.response
              : typeof payload?.text === "string"
                ? payload.text
                : "";
      if (rawAnswer) {
        if (live) {
          live.classList.remove("thinking", "scaffold");
        }
        const { cleanText, citations } = extractCitationBlock(rawAnswer);
        const sanitizedText = sanitizeAnswerText(cleanText, { allowExcerpts: excerptsAllowed });
        const normalizedText = normalizeAssistantText(sanitizedText || rawAnswer);
        const citeMap = buildCitationMap(citations);
        const normalizedSources = normalizeCitedSources([], citeMap);
        renderFinalAssistant(live, normalizedText, citeMap, {
          sources: normalizedSources,
          preserveCitationMap: true,
        });
        messageState.set(msgId, { streamText: rawAnswer, lastAnswer: normalizedText });
        logMessage("assistant", normalizedText, {
          model,
          id: msgId,
          requestId: payload?.requestId,
          sources: normalizedSources,
          citations: serializeCitationMap(citeMap),
          renderedMarkdown: normalizedText,
        });
        return;
      }
    }
    if (!response.body) throw new Error("Streaming response missing body");
    const reader = response.body.getReader();
    activeOwenStreamReader = reader;
    const decoder = new TextDecoder();

    let buffer = "";
    const getDeltaText = (evt) => {
      if (Array.isArray(evt.choices)) {
        const collected = evt.choices
          .map(choice => textFromParts(choice?.delta?.content || choice?.delta?.text))
          .filter(Boolean)
          .join("");
        if (collected) return collected;
      }
      return textFromParts(evt.delta ?? evt.output_text ?? evt.response?.output_text);
    };
    const flushBuffer = () => {
      throwIfAborted(signal);
      let idx;
      while ((idx = buffer.indexOf("\n\n")) !== -1) {
        const block = buffer.slice(0, idx);
        buffer = buffer.slice(idx + 2);
        const dataLine = block.split("\n").find(line => line.startsWith("data:"));
        if (!dataLine) continue;
        const payload = dataLine.slice(5).trim();
        if (!payload || payload === "[DONE]") continue;
        try {
          const evt = JSON.parse(payload);
          if (evt.error || evt.type === "response.error") {
            throw new Error(evt.error?.message || evt.error || "Chat error");
          }
          collectCitations(evt, citationEvents);
          const delta = getDeltaText(evt);
          if (delta) {
            if (!hasContent) {
              live.classList.remove("thinking");
              setBubbleStatus(live, "Answering…");
              startTypewriter(live, acc, { preserveStatus: true });
              hasContent = true;
            }
            acc += delta;
            messageState.set(msgId, { streamText: acc, lastAnswer: acc });
            updateTypewriter(live, acc);
            upsertConversationMessage({
              id: msgId,
              role: "assistant",
              text: acc,
              createdAt: Date.now(),
              model,
              docId: live?.dataset?.docId,
              requestId: live?.dataset?.requestId,
            });
            schedulePersistConversation();
            chatLog.scrollTop = chatLog.scrollHeight;
          }
        } catch (err) {
          console.error(err);
        }
      }
    };

    while (true) {
      throwIfAborted(signal);
      const { done, value } = await reader.read();
      if (value) buffer += decoder.decode(value, { stream: true });
      flushBuffer();
      if (done) break;
    }
    throwIfAborted(signal);
    buffer += decoder.decode(new Uint8Array(), { stream: false });
    flushBuffer();
    if (!acc) {
      const emptyTarget = ensureBubbleTextNode(live, { reset: true });
      emptyTarget.textContent = "Owen did not return any text.";
      messageState.set(msgId, { streamText: "", lastAnswer: "" });
    } else {
      const { cleanText, citations } = extractCitationBlock(acc);
      const sanitizedText = sanitizeAnswerText(cleanText, { allowExcerpts: excerptsAllowed });
      const normalizedText = normalizeAssistantText(sanitizedText || "Owen did not return any text.");
      const citeMap = buildCitationMap([...citationEvents, ...citations]);
      const normalizedSources = normalizeCitedSources([], citeMap);
      const storedCitations = serializeCitationMap(citeMap);
      messageState.set(msgId, { streamText: acc, lastAnswer: normalizedText });
      if (DEBUG_STREAMING) {
        console.log("[FINALIZE_CHAT]", {
          msgId,
          streamedLen: acc.length,
          finalAnswerLen: normalizedText.length,
          citationsCount: Object.keys(citeMap || {}).length,
        });
      }
      finishTypewriter(live);
      renderFinalAssistant(live, normalizedText, citeMap, {
        sources: normalizedSources,
        preserveCitationMap: true,
      });
      logMessage("assistant", normalizedText, {
        model,
        id: msgId,
        docId: live?.dataset?.docId,
        requestId: live?.dataset?.requestId,
        docTitle: activeConversationMeta.selectedDocTitle,
        sources: normalizedSources,
        citations: storedCitations,
        renderedMarkdown: normalizedText,
      });
    }
    if (DEBUG_STREAMING) {
      const state = messageState.get(msgId);
      console.log("[STREAM_FINAL]", {
        msgId,
        accLen: acc.length,
        stateStream: state?.streamText?.length || 0,
        stateLast: state?.lastAnswer?.length || 0,
      });
    }
  } catch (error) {
    if (isAbortError(error)) {
      if (live && live.isConnected) clearThinkingBubble(live);
      return;
    }
    console.error(error);
    if (live && live.isConnected) live.remove();
    const bubble = appendChatMessage("assistant", error instanceof Error ? error.message : "Something went wrong while chatting.");
    bubble.classList.add("error");
  } finally {
    clearOwenAbortController(abortController);
    setChatBusy(false);
  }
}

const appRoot = document.getElementById("app-root");
const ROUTE_ALIASES = {
  "/app": "/student",
  "/faculty/login": "/faculty",
};
const FACULTY_AUTH_KEY = "owen.facultyAuthed";
const FACULTY_AUTH_ERROR_KEY = "owen.facultyAuthError";
let globalListenersBound = false;
let studentMenuListenersBound = false;

function isFacultyAuthed() {
  try {
    return sessionStorage.getItem(FACULTY_AUTH_KEY) == "true";
  } catch {
    return false;
  }
}

function setFacultyAuthed(isAuthed) {
  try {
    sessionStorage.setItem(FACULTY_AUTH_KEY, isAuthed ? "true" : "false");
  } catch {}
}

function setFacultyAuthError(message) {
  try {
    if (message) {
      sessionStorage.setItem(FACULTY_AUTH_ERROR_KEY, message);
    } else {
      sessionStorage.removeItem(FACULTY_AUTH_ERROR_KEY);
    }
  } catch {}
}

function consumeFacultyAuthError() {
  try {
    const message = sessionStorage.getItem(FACULTY_AUTH_ERROR_KEY) || "";
    if (message) {
      sessionStorage.removeItem(FACULTY_AUTH_ERROR_KEY);
    }
    return message;
  } catch {
    return "";
  }
}

function handleFacultyUnauthorized(message = "Session expired - please log in again.") {
  setFacultyAuthed(false);
  setFacultyAuthError(message);
  if (activeRoute !== "faculty-login") {
    navigate("/faculty", { replace: true });
  }
}

async function facultyFetch(url, options = {}, label = "faculty") {
  const headers = new Headers(options.headers || {});
  const hasAuthHeader = headers.has("authorization");
  const request = { ...options, headers, credentials: "include" };
  console.info("[FACULTY] request", {
    label,
    url,
    credentials: request.credentials === "include",
    authorizationHeader: hasAuthHeader,
  });
  const resp = await fetch(url, request);
  if (resp.status === 401) {
    handleFacultyUnauthorized();
  }
  return resp;
}

async function ensureFacultySession() {
  try {
    const resp = await fetch("/api/faculty/session", { credentials: "include" });
    if (resp.status === 401) {
      handleFacultyUnauthorized();
      return false;
    }
    return resp.ok;
  } catch (error) {
    console.warn("[FACULTY] Session check failed", error);
    return false;
  }
}

function toggleTheme() {
  const next = document.body.classList.contains("theme-dark") ? "light" : "dark";
  applyTheme(next);
  localStorage.setItem(THEME_KEY, next);
}

function bindPanelToggle(panel, toggle, storageKey) {
  if (!panel || !toggle) return;
  const collapsed = safeGetItem(storageKey) == "true";
  if (collapsed) panel.classList.add("is-collapsed");
  toggle.setAttribute("aria-expanded", String(!collapsed));
  toggle.addEventListener("click", () => {
    const next = !panel.classList.contains("is-collapsed");
    panel.classList.toggle("is-collapsed", next);
    toggle.setAttribute("aria-expanded", String(!next));
    safeSetItem(storageKey, String(next));
  });
}

function initDetailsToggles(root) {
  root.querySelectorAll("[data-details]").forEach(btn => {
    const key = btn.getAttribute("data-details");
    const body = root.querySelector(`[data-details-body="${key}"]`);
    if (!body) return;
    btn.setAttribute("aria-expanded", "false");
    btn.addEventListener("click", () => {
      const isHidden = body.hasAttribute("hidden");
      body.toggleAttribute("hidden", !isHidden);
      btn.setAttribute("aria-expanded", String(isHidden));
    });
  });
}

function initLandingUI(root) {
  bindCommonElements(root);
  document.title = "OWEN";
  const container = root.querySelector('[data-role="hover-robot"]');
  if (container) {
    const robot = createHoverSwapRobot({
      staticSrc: "/android-chrome-512x512.png",
      animatedSrc: "/z_Upload-Image---Internal-Only-Style-dcbc2804.mp4",
      alt: "OWEN mascot",
    });
    container.replaceChildren(robot);
  }
  root.querySelectorAll("[data-route]").forEach(btn => {
    btn.addEventListener("click", () => {
      const path = btn.getAttribute("data-route") || "/";
      navigate(path);
    });
  });
}

function initStudentUI(root) {
  bindStudentElements(root);
  document.title = "OWEN Student Workspace";
  applyTheme(localStorage.getItem(THEME_KEY) == "dark" ? "dark" : "light");
  if (themeToggle) themeToggle.addEventListener("click", toggleTheme);
  bindComposerMetrics();
  setToolsDrawerOpen(false);
  closeQuiz();

  if (sidebarDrawerToggle) {
    sidebarDrawerToggle.addEventListener("click", () => toggleToolsDrawer());
  }
  if (toolsToggle) {
    toolsToggle.addEventListener("click", () => toggleToolsDrawer());
  }
  if (toolsOverlay) {
    toolsOverlay.addEventListener("click", () => setToolsDrawerOpen(false));
  }
  if (!toolsDrawerListenersBound) {
    toolsDrawerListenersBound = true;
    document.addEventListener("keydown", (event) => {
      if (event.key !== "Escape") return;
      if (!document.body.classList.contains("tools-open")) return;
      setToolsDrawerOpen(false);
    });
    if (typeof window.matchMedia === "function") {
      const media = window.matchMedia("(max-width: 1024px)");
      const syncDrawerState = () => setToolsDrawerOpen(document.body.classList.contains("tools-open"));
      if (typeof media.addEventListener === "function") {
        media.addEventListener("change", syncDrawerState);
      } else if (typeof media.addListener === "function") {
        media.addListener(syncDrawerState);
      }
    }
  }

  applySidebarCollapsed(safeGetItem(SIDEBAR_COLLAPSE_KEY) == "true");
  if (sidebarCollapseToggle) {
    sidebarCollapseToggle.addEventListener("click", () => {
      const next = !sidebar?.classList.contains("is-collapsed");
      applySidebarCollapsed(next);
    });
  }
  root.querySelectorAll("[data-sidebar-target]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetId = btn.getAttribute("data-sidebar-target");
      if (sidebar?.classList.contains("is-collapsed")) {
        applySidebarCollapsed(false);
      }
      if (!targetId) return;
      const target = root.querySelector(`#${targetId}`);
      if (target && target.scrollIntoView) {
        target.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    });
  });

  bindPanelToggle(historyPanel, historyToggle, HISTORY_COLLAPSE_KEY);
  bindPanelToggle(libraryPanel, libraryToggle, LIBRARY_COLLAPSE_KEY);
  bindPanelToggle(retrievePanel, retrieveToggle, RETRIEVE_COLLAPSE_KEY);

  if (libraryAdvancedToggle) {
    libraryAdvancedToggle.addEventListener("click", () => {
      setLibraryAdvancedState(!libraryAdvancedOpen);
    });
  }
  setLibraryAdvancedState(false);

  if (conversationSearch) {
    conversationSearch.addEventListener("input", renderConversationList);
  }

  if (newConvoBtn) newConvoBtn.addEventListener("click", () => startNewConversation());

  if (chatForm) {
    chatForm.addEventListener("submit", (event) => {
      event.preventDefault();
      if (chatStreaming) return;
      const prompt = chatInput.value.trim();
      if (!prompt) {
        chatInput.focus();
        return;
      }
      runChat(prompt);
    });
  }

  if (chatInput) {
    chatInput.addEventListener("keydown", (event) => {
      if (event.key == "Enter" && !event.shiftKey) {
        event.preventDefault();
        if (chatStreaming) return;
        const prompt = chatInput.value.trim();
        if (prompt) runChat(prompt);
      }
    });
    chatInput.addEventListener("input", () => updateComposerMetrics());
    chatInput.addEventListener("paste", handleChatPaste);
  }

  if (chatLog) {
    chatLog.addEventListener("click", (event) => {
      if ((event.target instanceof Element) && event.target.closest(".bubble.assistant")) {
        skipAllTypewriters();
      }
    });
    chatLog.addEventListener("wheel", () => skipAllTypewriters(), { passive: true });
    chatLog.addEventListener("scroll", () => hideCitationTooltip(), { passive: true });
  }

  if (chatAttach && chatFileInput) {
    chatAttach.addEventListener("click", () => {
      chatFileInput.value = "";
      chatFileInput.click();
    });
  }

  if (chatFileInput) {
    chatFileInput.addEventListener("change", async () => {
      const file = chatFileInput.files?.[0];
      if (!file) return;
      activePdfSession = null;
      clearOcrSession();
      attachments.length = 0;
      if (oneShotPreviewUrl && oneShotPreviewUrl.startsWith("blob:")) URL.revokeObjectURL(oneShotPreviewUrl);
      oneShotFile = file;
      if (file.type.startsWith("image/")) {
        const preview = await generatePreviewDataUrl(file);
        oneShotPreviewUrl = preview || "";
      } else {
        oneShotPreviewUrl = "";
      }
      updateAttachmentIndicator();
    });
  }

  if (retrieveBtn) {
    retrieveBtn.addEventListener("click", retrieveByCode);
  }
  if (retrieveInput) {
    retrieveInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        retrieveByCode();
      }
    });
  }
  if (retrieveTypeSelect) {
    retrieveTypeSelect.addEventListener("change", () => {
      renderRetrieveStudyGuide(null);
      renderRetrieveAnki(null);
      setStatus(retrieveStatus, "info", "");
    });
  }

  if (libraryRefreshBtn) libraryRefreshBtn.addEventListener("click", () => loadLibraryList());
  if (librarySelect) librarySelect.addEventListener("change", onLibrarySelectChange);
  if (libraryActionBtn) libraryActionBtn.addEventListener("click", onLibraryAction);
  if (libraryQuizBtn) libraryQuizBtn.addEventListener("click", onLibraryQuizAction);
  if (libraryClearBtn) libraryClearBtn.addEventListener("click", () => {
    setLibrarySelection(null);
    if (librarySelect) librarySelect.value = "";
  });
  if (libraryIndexBtn) libraryIndexBtn.addEventListener("click", runLibraryBatchIndex);
  if (libraryIngestBtn) libraryIngestBtn.addEventListener("click", (event) => {
    const mode = event.shiftKey || event.metaKey ? "full" : "embedded_only";
    runLibraryBatchIngest(mode);
  });

  if (sourcesDrawerClose) sourcesDrawerClose.addEventListener("click", closeSourcesDrawer);
  if (sourcesOverlay) sourcesOverlay.addEventListener("click", closeSourcesDrawer);
  if (sourcesSearch) {
    sourcesSearch.addEventListener("input", () => {
      if (activeSourcesCard) renderSourcesDrawer(activeSourcesCard);
    });
  }
  if (quizCloseBtn) quizCloseBtn.addEventListener("click", closeQuiz);
  if (quizOverlay) quizOverlay.addEventListener("click", closeQuiz);
  if (quizModal) quizModal.addEventListener("click", handleQuizOptionClick);
  if (quizInterruptBtn) {
    quizInterruptBtn.addEventListener("click", () => interruptQuizSession());
  }
  if (quizRetryBtn) {
    quizRetryBtn.addEventListener("click", () => {
      runQuizBatch();
    });
  }
  if (quizGenerateMoreBtn) {
    quizGenerateMoreBtn.addEventListener("click", () => runQuizBatch());
  }
  if (!quizEscapeListenerBound) {
    quizEscapeListenerBound = true;
    document.addEventListener("keydown", (event) => {
      if (event.key !== "Escape") return;
      if (!document.body.classList.contains("quiz-open")) return;
      closeQuiz();
    });
  }

  if (!studentMenuListenersBound) {
    studentMenuListenersBound = true;
    document.addEventListener("click", (event) => {
      const target = event.target instanceof Element ? event.target : null;
      if (!target) return;
      if (target.closest(".menu-popover") || target.closest(".menu-trigger")) return;
      if (openConversationMenuId || openExportMenuId || pendingDeleteId) {
        openConversationMenuId = null;
        openExportMenuId = null;
        pendingDeleteId = null;
        renderConversationList();
      }
    });
  }

  updateAttachmentIndicator();
  renderConversationList();
  const lastSavedId = loadActiveConversationId();
  if (lastSavedId) {
    loadConversation(lastSavedId);
  } else if (storedConversations.length) {
    loadConversation(storedConversations[0].id);
  }
  syncConversationIndexFromServer({
    force: true,
    bootstrap: !lastSavedId && !storedConversations.length,
  });
  setTimeout(() => backfillLocalConversationsToServer(), 1200);
  setLibrarySelection(null);
  loadLibraryList().catch(() => {});
}

function initFacultyLoginUI(root) {
  bindFacultyElements(root);
  document.title = "OWEN Faculty Login";
  applyTheme(localStorage.getItem(THEME_KEY) == "dark" ? "dark" : "light");
  if (themeToggle) themeToggle.addEventListener("click", toggleTheme);
  const showError = (message) => {
    if (!facultyLoginError) return;
    if (!message) {
      facultyLoginError.hidden = true;
      facultyLoginError.textContent = "";
      return;
    }
    facultyLoginError.hidden = false;
    facultyLoginError.textContent = message;
  };
  const pendingError = consumeFacultyAuthError();
  if (pendingError) showError(pendingError);
  if (isFacultyAuthed()) {
    ensureFacultySession().then((ok) => {
      if (ok) {
        navigate("/faculty/pipeline", { replace: true });
        return;
      }
      const staleError = consumeFacultyAuthError();
      if (staleError) showError(staleError);
    });
  }
  const attemptUnlock = async () => {
    const input = facultyPasscodeInput?.value.trim() || "";
    if (!input) {
      showError("Enter the faculty passcode.");
      return;
    }
    if (facultyUnlockBtn) facultyUnlockBtn.disabled = true;
    try {
      const resp = await fetch("/api/faculty/login", {
        method: "POST",
        headers: { "content-type": "application/json" },
        credentials: "include",
        body: JSON.stringify({ passcode: input }),
      });
      const { payload, rawText } = await readJsonResponse(resp);
      if (!resp.ok || payload?.error) {
        const msg = extractErrorMessage(payload, rawText, "Login failed.");
        throw new Error(msg);
      }
      setFacultyAuthed(true);
      setFacultyAuthError("");
      showError("");
      navigate("/faculty/pipeline");
    } catch (error) {
      showError(error instanceof Error ? error.message : "Login failed.");
    } finally {
      if (facultyUnlockBtn) facultyUnlockBtn.disabled = false;
    }
  };
  if (facultyUnlockBtn) facultyUnlockBtn.addEventListener("click", attemptUnlock);
  if (facultyPasscodeInput) {
    facultyPasscodeInput.addEventListener("keydown", (event) => {
      if (event.key == "Enter") {
        event.preventDefault();
        attemptUnlock();
      }
    });
    facultyPasscodeInput.addEventListener("input", () => showError(""));
  }
}

function initFacultyPipelineUI(root) {
  bindFacultyElements(root);
  libraryPendingDeleteId = null;
  document.title = "OWEN Faculty Pipeline";
  applyTheme(localStorage.getItem(THEME_KEY) == "dark" ? "dark" : "light");
  if (themeToggle) themeToggle.addEventListener("click", toggleTheme);
  void ensureFacultySession();
  if (!machineStudyBtn) console.warn("[FACULTY] Missing generate study guide button.");
  if (!publishStudyGuideBtn) console.warn("[FACULTY] Missing publish study guide button.");
  if (facultyLogoutBtn) {
    facultyLogoutBtn.addEventListener("click", async () => {
      try {
        await fetch("/api/faculty/logout", { method: "POST", credentials: "include" });
      } catch (error) {
        console.warn("[FACULTY] Logout failed", error);
      }
      setFacultyAuthed(false);
      setFacultyAuthError("");
      navigate("/faculty", { replace: true });
    });
  }
  initDetailsToggles(root);
  initBucketSelect();
  requestAnimationFrame(() => syncPipelineDualCardHeight(root));

  if (fullDocToggle) {
    fullDocToggle.checked = true;
    fullDocToggle.addEventListener("change", () => {
      fullDocumentMode = Boolean(fullDocToggle.checked);
    });
  }

  if (fileButton && fileInput) fileButton.addEventListener("click", () => fileInput.click());
  if (fileInput) {
    fileInput.addEventListener("change", () => {
      const file = fileInput.files?.[0];
      if (fileNameDisplay) fileNameDisplay.textContent = file ? file.name : "No file selected";
      updateSendButtonState();
    });
  }
  if (nameInput) nameInput.addEventListener("input", updateSendButtonState);
  if (sendButton) {
    sendButton.textContent = "Upload lecture";
    sendButton.addEventListener("click", uploadFile);
  }

  if (librarySelect) librarySelect.addEventListener("change", onLibrarySelectChange);
  if (libraryActionBtn) libraryActionBtn.addEventListener("click", onLibraryAction);
  if (libraryIndexBtn) libraryIndexBtn.addEventListener("click", runLibraryBatchIndex);
  if (libraryIngestBtn) libraryIngestBtn.addEventListener("click", (event) => {
    const mode = event.shiftKey || event.metaKey ? "full" : "embedded_only";
    runLibraryBatchIngest(mode);
  });
  if (libraryManagerSearch) {
    libraryManagerSearch.addEventListener("input", () => renderLibraryManagerList());
  }
  if (libraryManagerList) {
    libraryManagerList.addEventListener("click", handleLibraryManagerClick);
    libraryManagerList.addEventListener("change", handleLibraryManagerChange);
    libraryManagerList.addEventListener("input", handleLibraryManagerInput);
  }

  if (machineLectureSelect) machineLectureSelect.addEventListener("change", updateMachineSelection);
  if (machineGenerateBtn) machineGenerateBtn.addEventListener("click", runMachineTxtGeneration);
  if (machineTxtButton && machineTxtInput) machineTxtButton.addEventListener("click", () => machineTxtInput.click());
  if (machineTxtInput) {
    machineTxtInput.addEventListener("change", () => {
      if (machineUseLastToggle) machineUseLastToggle.checked = false;
      if (machineStudyResult) machineStudyResult.innerHTML = "";
      updateMachineStudyState();
    });
  }
  if (machineUseLastToggle) {
    machineUseLastToggle.addEventListener("change", () => {
      if (machineUseLastToggle.checked && machineTxtInput) {
        machineTxtInput.value = "";
      }
      if (machineStudyResult) machineStudyResult.innerHTML = "";
      updateMachineStudyState();
    });
  }
  bindOnce(machineStudyBtn, "click", runMachineStudyGuideGeneration, "study-guide-generate");
  bindOnce(publishStudyGuideBtn, "click", runStudyGuidePublish, "study-guide-publish");

  if (ankiPdfButton && ankiPdfInput) {
    ankiPdfButton.addEventListener("click", () => {
      ankiPdfInput.value = "";
      ankiPdfInput.click();
    });
  }
  if (ankiPdfInput) {
    ankiPdfInput.addEventListener("change", async () => {
      const file = ankiPdfInput.files?.[0] || null;
      const rejectPdf = (message) => {
        ankiPdfFile = null;
        ankiPdfConverting = false;
        ankiImagesSource = "";
        ankiImageFiles = [];
        if (ankiPdfInput) ankiPdfInput.value = "";
        clearAnkiResult();
        setAnkiPdfStatus("error", message);
        clearAnkiDerivedArtifacts();
        setAnkiJobError(message, "Upload a valid PDF and try again.");
        updateAnkiGenerateState({ preserveStatus: true });
      };
      if (!file) return;
      console.log("[ANKI] PDF selected", { name: file.name, type: file.type, size: file.size });
      if (!file.size) {
        rejectPdf("PDF file rejected: file is empty.");
        return;
      }
      if (!isAnkiPdfFile(file)) {
        const message = buildAnkiPdfError(file);
        rejectPdf(message);
        return;
      }
      clearAnkiGenerateError();
      clearAnkiDerivedArtifacts();
      clearAnkiResult();
      ankiPdfFile = file;
      ankiPdfConverting = false;
      ankiImagesSource = "";
      ankiImageFiles = [];
      ankiTranscriptFile = null;
      ankiBoldmapFile = null;
      ankiClassmateFile = null;
      if (ankiTranscriptInput) ankiTranscriptInput.value = "";
      if (ankiBoldmapInput) ankiBoldmapInput.value = "";
      if (ankiClassmateInput) ankiClassmateInput.value = "";
      if (ankiImagesInput) ankiImagesInput.value = "";
      setAnkiPdfStatus("info", "PDF loaded. Click Parse to convert slides.");
      updateAnkiGenerateState();
    });
  }
  if (ankiParseButton) {
    ankiParseButton.addEventListener("click", () => {
      runAnkiParse();
    });
  }

  if (ankiImagesButton && ankiImagesInput) {
    ankiImagesButton.addEventListener("click", () => {
      ankiImagesInput.value = "";
      ankiImagesInput.click();
    });
  }
  if (ankiImagesInput) {
    ankiImagesInput.addEventListener("change", async () => {
      const files = Array.from(ankiImagesInput.files || []);
      if (!files.length) return;
      const invalidFiles = files.filter(file => !isAnkiImageFile(file));
      if (invalidFiles.length) {
        const message = buildAnkiImagesError(invalidFiles[0]);
        setAnkiJobError(message, "Upload JPG/PNG images and try again.");
        ankiImagesInput.value = "";
        ankiImageFiles = [];
        ankiImagesSource = "";
        clearAnkiResult();
        updateAnkiGenerateState({ preserveStatus: true });
        return;
      }
      clearAnkiGenerateError();
      if (files.length > ANKI_MAX_IMAGES) {
        setAnkiJobError(`Too many images. Max is ${ANKI_MAX_IMAGES}.`, "Upload fewer images and try again.");
        ankiImagesInput.value = "";
        return;
      }
      clearAnkiDerivedArtifacts();
      clearAnkiResult();
      ankiPdfFile = null;
      ankiTranscriptFile = null;
      ankiBoldmapFile = null;
      ankiClassmateFile = null;
      if (ankiPdfInput) ankiPdfInput.value = "";
      if (ankiTranscriptInput) ankiTranscriptInput.value = "";
      if (ankiBoldmapInput) ankiBoldmapInput.value = "";
      if (ankiClassmateInput) ankiClassmateInput.value = "";
      ankiPdfConverting = true;
      updateAnkiGenerateState({ preserveStatus: true });
      if (ankiGenerateStatus) {
        setStatus(ankiGenerateStatus, "pending", "Preparing images...");
      }
      try {
        const normalized = await normalizeAnkiImageFiles(files);
        ankiImageFiles = normalized;
        ankiImagesSource = "upload";
        await buildAnkiDerivedArtifacts({
          images: normalized,
          fileHash: createAnkiRequestId(),
          ocrMode: "background",
        });
        setAnkiPdfStatus("success", `${normalized.length} images ready.`);
      } catch (error) {
        const message = error instanceof Error ? error.message : "Image preparation failed.";
        ankiImageFiles = [];
        ankiImagesSource = "";
        setAnkiJobError(message, "Upload JPG/PNG images and try again.");
        setAnkiPdfStatus("error", message);
      } finally {
        ankiPdfConverting = false;
        updateAnkiGenerateState();
      }
    });
  }

  if (ankiTranscriptButton && ankiTranscriptInput) {
    ankiTranscriptButton.addEventListener("click", () => {
      ankiTranscriptInput.value = "";
      ankiTranscriptInput.click();
    });
  }
  if (ankiTranscriptInput) {
    ankiTranscriptInput.addEventListener("change", () => {
      const file = ankiTranscriptInput.files?.[0] || null;
      if (!file) return;
      if (!ankiImageFiles.length) {
        ankiTranscriptInput.value = "";
        setAnkiJobError("Load slides before uploading a transcript.", "Parse a PDF or upload JPEGs first.");
        updateAnkiGenerateState({ preserveStatus: true });
        return;
      }
      if (!isAnkiTranscriptFile(file)) {
        const message = buildAnkiTranscriptError(file);
        ankiTranscriptFile = null;
        ankiTranscriptInput.value = "";
        clearAnkiResult();
        setAnkiJobError(message, "Upload a TXT transcript and try again.");
        updateAnkiGenerateState({ preserveStatus: true });
        return;
      }
      clearAnkiGenerateError();
      ankiTranscriptFile = file;
      clearAnkiResult();
      updateAnkiGenerateState();
    });
  }

  if (ankiBoldmapButton && ankiBoldmapInput) {
    ankiBoldmapButton.addEventListener("click", () => {
      ankiBoldmapInput.value = "";
      ankiBoldmapInput.click();
    });
  }
  if (ankiBoldmapInput) {
    ankiBoldmapInput.addEventListener("change", () => {
      ankiBoldmapFile = ankiBoldmapInput.files?.[0] || null;
      clearAnkiGenerateError();
      clearAnkiResult();
      updateAnkiGenerateState();
    });
  }

  if (ankiClassmateButton && ankiClassmateInput) {
    ankiClassmateButton.addEventListener("click", () => {
      ankiClassmateInput.value = "";
      ankiClassmateInput.click();
    });
  }
  if (ankiClassmateInput) {
    ankiClassmateInput.addEventListener("change", () => {
      ankiClassmateFile = ankiClassmateInput.files?.[0] || null;
      clearAnkiGenerateError();
      clearAnkiResult();
      updateAnkiGenerateState();
    });
  }

  if (ankiLoadedList) {
    ankiLoadedList.addEventListener("click", (event) => {
      const target = event.target instanceof Element ? event.target.closest("[data-anki-clear]") : null;
      if (!target) return;
      const key = target.getAttribute("data-anki-clear") || "";
      if (key) handleAnkiClear(key);
    });
  }
  bindOnce(ankiGenerateBtn, "click", runAnkiGeneration, "anki-generate");
  ankiJobState = ANKI_JOB_STATES.EMPTY;
  ankiLastStableState = ANKI_JOB_STATES.EMPTY;
  setAnkiPdfStatus("", "");
  updateAnkiGenerateState();

  metaDataUnlocked = true;
  mountMetaBody();
  resetMetaPanel();
  applyMetaExpandedState(true);
  updateMetaLectureOptions();
  refreshMetaDataPanel();

  updateSendButtonState();
  hydrateStudyGuideState();
  updateMachineStudyState();
  updatePublishState();
  setLibrarySelection(null);
  loadLibraryList().catch(() => {});
}

function initGlobalListeners() {
  if (globalListenersBound) return;
  globalListenersBound = true;

  document.addEventListener("focusin", (event) => {
    const target = event.target instanceof Element ? event.target : null;
    const chip = target?.closest(".citation-chip");
    if (chip) {
      showCitationTooltip(chip);
      return;
    }
    if (citationTooltip && target && citationTooltip.contains(target)) {
      if (tooltipHideTimer) {
        clearTimeout(tooltipHideTimer);
        tooltipHideTimer = null;
      }
    }
  });

  document.addEventListener("focusout", (event) => {
    const related = event.relatedTarget instanceof Element ? event.relatedTarget : null;
    if (related && ((related.closest(".citation-chip")) || (citationTooltip && citationTooltip.contains(related)))) return;
    scheduleTooltipHide();
  });

  document.addEventListener("pointerdown", (event) => {
    const target = event.target instanceof Element ? event.target : null;
    if (!target) return;
    if (target.closest(".citation-chip")) return;
    if (citationTooltip && citationTooltip.contains(target)) return;
    hideCitationTooltip();
  });

  document.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target : null;
    const chip = target?.closest(".citation-chip");
    if (!chip || chip.closest(".citation-tooltip")) return;
    if (chip.dataset.citeDisabled == "true" || chip.getAttribute("aria-disabled") == "true") return;
    const info = getCitationInfoFromChip(chip);
    if (info?.card && info?.id && info?.info) {
      openSourcesDrawer(info.card, info.id);
    }
  });

  document.addEventListener("keydown", handleDrawerKeydown);
  document.addEventListener("keydown", (event) => {
    if (event.key == "Escape") hideCitationTooltip();
    if (activeRoute == "student" && (event.key == " " || event.key == "Enter")) {
      skipAllTypewriters();
    }
  });

  document.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target.closest("a.citation-index") : null;
    if (!target) {
      return;
    }
    const href = target.getAttribute("href") || "";
    if (!href.startsWith("#")) {
      return;
    }
    event.preventDefault();
    const referenceNode = document.querySelector(href);
    if (referenceNode) {
      referenceNode.classList.add("ref-highlight");
      referenceNode.scrollIntoView({ block: "center", behavior: "smooth" });
      setTimeout(() => referenceNode.classList.remove("ref-highlight"), 800);
    }
  });

  window.addEventListener("hashchange", () => {
    scrollToHash({ behavior: "smooth" });
  });

  window.addEventListener("resize", () => {
    if (activeRoute === "faculty-pipeline") {
      syncPipelineDualCardHeight();
    }
    if (activeRoute === "student") {
      updateComposerMetrics();
    }
  });

  if (window.visualViewport) {
    window.visualViewport.addEventListener("resize", () => {
      if (activeRoute === "student") {
        updateComposerMetrics();
      }
    });
  }
}

function normalizePath(pathname) {
  if (!pathname) return "/";
  let path = pathname;
  if (path.length > 1 && path.endsWith("/")) {
    path = path.slice(0, -1);
  }
  return path || "/";
}

function resolveRoute(pathname) {
  const normalized = normalizePath(pathname);
  const aliased = ROUTE_ALIASES[normalized] || normalized;
  if (aliased == "/student") return { route: "student", path: "/student" };
  if (aliased == "/faculty/pipeline") return { route: "faculty-pipeline", path: "/faculty/pipeline" };
  if (aliased.startsWith("/faculty")) return { route: "faculty-login", path: "/faculty" };
  if (aliased == "/") return { route: "landing", path: "/" };
  return { route: "landing", path: "/" };
}

function renderRoute(pathname) {
  if (!appRoot) return;
  const resolved = resolveRoute(pathname);
  let route = resolved.route;
  let path = resolved.path;
  if (route == "faculty-pipeline" && !isFacultyAuthed()) {
    route = "faculty-login";
    path = "/faculty";
  }
  if (path != window.location.pathname) {
    history.replaceState({}, "", path);
  }
  activeRoute = route;
  document.body.dataset.route = route == "landing" ? "landing" : route.startsWith("faculty") ? "faculty" : "student";
  document.body.classList.remove("tools-open");
  resetBoundElements();
  appRoot.innerHTML = "";
  const templateId = route == "landing"
    ? "landing-template"
    : route == "student"
      ? "student-template"
      : route == "faculty-pipeline"
        ? "faculty-pipeline-template"
        : "faculty-login-template";
  const template = document.getElementById(templateId);
  if (!template) return;
  appRoot.appendChild(template.content.cloneNode(true));
  if (route == "landing") initLandingUI(appRoot);
  if (route == "student") initStudentUI(appRoot);
  if (route == "faculty-login") initFacultyLoginUI(appRoot);
  if (route == "faculty-pipeline") initFacultyPipelineUI(appRoot);
  initGlobalListeners();
}

function navigate(path, opts = {}) {
  const replace = Boolean(opts.replace);
  if (replace) {
    history.replaceState({}, "", path);
  } else {
    history.pushState({}, "", path);
  }
  renderRoute(path);
}

function bootstrap() {
  renderRoute(window.location.pathname);
  window.addEventListener("popstate", () => renderRoute(window.location.pathname));
}

bootstrap();

export {
  decorateAssistantResponse,
  applyResponseView,
  finalizeResponseCard,
  getResponseState,
  setAllSectionsOpen,
  buildCitationMapFromStored,
};
