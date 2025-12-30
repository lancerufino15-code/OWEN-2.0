/**
 * Legacy chat UI controller (browser script).
 *
 * Used by: root `index.html` via `<script type="module" src="/chat.js">`.
 *
 * Responsibilities:
 * - Bind UI elements, handle chat input/streaming, uploads, retrieval, and history.
 *
 * Assumptions:
 * - Runs in a browser with DOM access; no server-side rendering.
 */
const fileInput = document.getElementById("file-input");
const fileButton = document.getElementById("file-button");
const fileNameDisplay = document.getElementById("file-name");
const nameInput = document.getElementById("name-input");
const bucketSelect = document.getElementById("bucket-select");
const sendButton = document.getElementById("send-button");
const uploadStatus = document.getElementById("upload-status");
const uploadPanel = document.getElementById("uploadPanel");
const uploadPanelBody = document.getElementById("uploadPanelBody");
const uploadToggle = document.getElementById("uploadToggle");
const uploadUnlockInput = document.getElementById("uploadUnlockInput");
const uploadLockBtn = document.getElementById("uploadLock");

const retrieveInput = document.getElementById("retrieve-input");
const retrieveBtn = document.getElementById("retrieve-btn");
const retrieveStatus = document.getElementById("retrieve-status");

const savedPlaceholder = document.getElementById("saved-placeholder");
const savedList = document.getElementById("saved-list");
const newConvoBtn = document.getElementById("new-convo");
const historyPanel = document.getElementById("historyPanel");
const historyToggle = document.getElementById("historyToggle");
const historyList = document.getElementById("historyList");

const chatLog = document.getElementById("chat-log");
const chatForm = document.getElementById("chat-form");
const chatInput = document.getElementById("chat-input");
const chatSend = document.getElementById("chat-send");
const chatAttach = document.getElementById("chat-attach");
const chatFileInput = document.getElementById("chat-file-input");
const attachStatus = document.getElementById("attach-status");
const insightsLog = document.getElementById("insights-log");
const hashtagCloud = document.getElementById("hashtag-cloud");
const metaTagsBox = document.getElementById("meta-tags-box");
const metaDataPanel = document.getElementById("metaDataPanel");
const metaUpdated = document.getElementById("metaUpdated");
const metaStatus = document.getElementById("metaStatus");
const metaLectureSelect = document.getElementById("metaLectureSelect");
const metaUnlockInput = document.getElementById("metaUnlockInput");
const metaLockBtn = document.getElementById("metaLockBtn");
const metaError = document.getElementById("metaError");
const modelButtons = Array.from(document.querySelectorAll(".model-btn"));
const themeToggle = document.getElementById("theme-toggle");
const fullDocToggle = document.getElementById("full-doc-toggle");
const librarySearchInput = document.getElementById("library-search");
const librarySearchBtn = document.getElementById("library-search-btn");
const libraryResults = document.getElementById("library-results");
const libraryStatus = document.getElementById("library-status");
const libraryIndexBtn = document.getElementById("library-index-btn");
const libraryIngestBtn = document.getElementById("library-ingest-btn");
const librarySelection = document.getElementById("library-selection");
const librarySelect = document.getElementById("librarySelect");
const libraryStatusBadge = document.getElementById("libraryStatusBadge");
const libraryActionBtn = document.getElementById("libraryActionBtn");
const libraryRefreshBtn = document.getElementById("libraryRefresh");
const libraryClearBtn = document.getElementById("libraryClear");

const THEME_TOGGLE_MOON_ICON = `
  <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path>
  </svg>
`.trim();

const THEME_TOGGLE_LIGHT_ICON = `
  <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M8 9.5a4 4 0 1 1 8 0c0 1.6-.7 2.8-1.8 3.8-.5.5-.8 1.1-1 1.7h-2.4c-.2-.6-.5-1.2-1-1.7C8.7 12.3 8 11.1 8 9.5z"></path>
    <path d="M9.5 18h5"></path>
    <path d="M10.5 21h3"></path>
  </svg>
`.trim();

const requiredNodes = [
  fileInput,
  fileButton,
  fileNameDisplay,
  nameInput,
  bucketSelect,
  sendButton,
  uploadStatus,
  uploadPanel,
  uploadPanelBody,
  uploadToggle,
  uploadUnlockInput,
  uploadLockBtn,
  retrieveInput,
  retrieveBtn,
  retrieveStatus,
  savedPlaceholder,
  savedList,
  newConvoBtn,
  historyPanel,
  historyToggle,
  historyList,
  chatLog,
  chatForm,
  chatInput,
  chatSend,
  chatAttach,
  chatFileInput,
  attachStatus,
  insightsLog,
  hashtagCloud,
  metaTagsBox,
  metaDataPanel,
  metaUpdated,
  metaStatus,
  metaLectureSelect,
  metaUnlockInput,
  metaLockBtn,
  metaError,
  themeToggle,
  libraryIndexBtn,
  libraryIngestBtn,
  librarySelect,
  libraryStatusBadge,
  libraryActionBtn,
  libraryRefreshBtn,
  libraryClearBtn,
  ...modelButtons,
];

if (requiredNodes.some(node => !node)) {
  throw new Error("OWEN UI is missing one or more required DOM nodes.");
}

const attachmentUrlMap = window.__attachmentUrlMap || (window.__attachmentUrlMap = {});
const attachmentSummaries = window.__attachmentSummaries || (window.__attachmentSummaries = {});
const documentSummaryTasks = window.__documentSummaryTasks || (window.__documentSummaryTasks = new Set());
const attachments = [];
const conversation = [];
let chatStreaming = false;
const TABLE_WRAP_PREF_KEY = "owen.tableWrapDefault";
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
const MAX_STORED_MESSAGES = 400;
const MAX_STORED_CHARS = 600_000;
const MAX_MESSAGE_CHARS = 120_000;
const TRUNCATION_NOTICE = " … (truncated for storage)";
const OCR_PAGE_LIMIT = 15;
const PDFJS_VERSION = "4.10.38";
const PDF_RENDER_MAX_DIMENSION = 1600;
const PDF_JPEG_QUALITY = 0.75;
const OCR_CONCURRENCY = 2;
const EARLY_OCR_READY_COUNT = 3;
const OCR_MAX_RETRIES = 3;
const OCR_SAVE_BATCH = 5;
const OCR_SESSION_KEY = "owen.ocr_session";
const DEST_STORAGE_KEY = "owen.uploadDestination";
const DESTINATIONS = [
  { label: "Anki Decks", value: "anki_decks" },
  { label: "Study Guides", value: "study_guides" },
  { label: "Library", value: "library" },
];
const CHAT_BUILD_ID = "20250222-1";
const TYPEWRITER_SPEED_MS = 28;
const TYPEWRITER_DEBOUNCE_MS = 90;
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
let activeConversationId = loadActiveConversationId() || generateConversationId();
let activeConversationMeta = {
  id: activeConversationId,
  title: "Conversation",
  createdAt: Date.now(),
  updatedAt: Date.now(),
  selectedDocId: null,
  selectedDocTitle: "",
};
const pendingMathNodes = [];
let katexIntervalStarted = false;
const metaTagSet = new Set();
let metaDataUnlocked = false;
let metaAnalyticsController = null;
let metaAnalyticsRequestId = 0;
let oneShotFile = null;
let oneShotPreviewUrl = "";
let activePdfSession = null;
let fullDocumentMode = true;
let activeLibraryDoc = null;
let libraryListItems = [];
let librarySearchResults = [];
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

function highlightKeywords(str) {
  return str.replace(/\b(important|key|note|tip|warning|result|action|next steps?)\b/gi, (match) => `<mark>${match}</mark>`);
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

function isLikelyValidCitationUrl(url) {
  if (!url) return false;
  if (url.startsWith("/")) return true; // allow in-app attachment links
  if (!/^https?:\/\//i.test(url)) return false;
  try {
    const parsed = new URL(url);
    const host = (parsed.hostname || "").toLowerCase();
    const blockedHosts = new Set([
      "example.com",
      "example.org",
      "example.net",
      "example.edu",
      "example1.com",
      "example2.com",
      "example3.com",
      "test.com",
      "localhost",
      "127.0.0.1",
      "0.0.0.0",
    ]);
    if (blockedHosts.has(host)) return false;
    if (/^127\./.test(host) || /^0\./.test(host)) return false;
    if (host.endsWith(".invalid") || host.endsWith(".local")) return false;
    if (!host.includes(".")) return false;
    if (/placeholder|dummy|lorem/i.test(url)) return false;
    return true;
  } catch {
    return false;
  }
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
      map.set(id, url);
    }
  });
  return map;
}

function enrichCitationMapFromText(text, map) {
  const refRegex = /(\d+)\.\s.*?(https?:\/\/\S+)/g;
  let match;
  while ((match = refRegex.exec(text))) {
    const id = match[1];
    let url = match[2];
    url = url.replace(/[\]\)]?[\.,;!?]*$/, "");
    if (isLikelyValidCitationUrl(url)) {
      map.set(id, url);
    }
  }
  parseCitationLines(text).forEach(({ id, url }) => {
    if (id && isLikelyValidCitationUrl(url)) map.set(String(id), url);
  });
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
  upsertConversationMessage(entry);
  persistConversation();
}

function persistConversation() {
  const now = Date.now();
  const title = conversation.length ? deriveConversationTitle(conversation) : (activeConversationMeta.title || "Conversation");
  const selectedDocId = activeLibraryDoc?.docId || activeConversationMeta.selectedDocId || null;
  const selectedDocTitle = activeLibraryDoc?.title || activeLibraryDoc?.key || activeConversationMeta.selectedDocTitle || "";
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
    const key = msg.id ? `id:${msg.id}` : `${msg.role}|${msg.text}|${msg.imageUrl || ""}|${(msg.attachments || []).length}`;
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
  storedConversations.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  savedList.innerHTML = "";
  if (!storedConversations.length) {
    savedPlaceholder.hidden = false;
    return;
  }
  savedPlaceholder.hidden = true;
  storedConversations.forEach(conv => {
    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.type = "button";
    const title = document.createElement("div");
    title.textContent = conv.title || "Conversation";
    title.style.fontWeight = "600";
    const meta = document.createElement("small");
    meta.textContent = new Date(conv.updatedAt || Date.now()).toLocaleString();
    btn.append(title, meta);
    btn.classList.toggle("active", conv.id === activeConversationId);
    btn.addEventListener("click", () => loadConversation(conv.id));
    const del = document.createElement("button");
    del.type = "button";
    del.className = "delete-btn";
    del.textContent = "×";
    del.title = "Delete conversation";
    del.setAttribute("aria-label", "Delete conversation");
    del.addEventListener("click", (event) => {
      event.stopPropagation();
      const idx = storedConversations.findIndex(c => c.id === conv.id);
      if (idx >= 0) {
        storedConversations.splice(idx, 1);
        deleteConversationDocument(conv.id);
        saveStoredConversations();
        if (activeConversationId === conv.id) {
          startNewConversation();
        } else {
          renderConversationList();
        }
      }
    });
    li.append(btn, del);
    savedList.append(li);
  });
}

async function loadConversation(id) {
  if (id && id !== activeConversationId) {
    persistConversation();
  }
  const entry = storedConversations.find(conv => conv.id === id);
  const doc = loadConversationDocument(id);
  if (!entry && !doc) return;
  activeConversationId = id;
  const loaded = doc || { id, messages: [], title: entry?.title || "Conversation", createdAt: Date.now(), updatedAt: Date.now(), selectedDocId: null, selectedDocTitle: "" };
  activeConversationMeta = {
    id,
    title: loaded.title || entry?.title || "Conversation",
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
    saveStoredConversations();
  }
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
    insightsLog.textContent = `Last ask: ${lastUser.text}`;
    updateHashtagCloud(lastUser.text);
  } else {
    insightsLog.textContent = "Conversation loaded.";
    hashtagCloud.innerHTML = "";
  }
  rebuildMetaTagsFromConversation(messages);
  if (activeConversationMeta.selectedDocId) {
    setLibrarySelection({
      docId: activeConversationMeta.selectedDocId,
      title: activeConversationMeta.selectedDocTitle || activeConversationMeta.selectedDocId,
      status: "ready",
    });
  } else {
    setLibrarySelection(null);
  }
  await syncMetaTagsFromKV(id);
  renderMathIfReady(chatLog);
  renderConversationList();
}

function startNewConversation() {
  if (conversation.length) persistConversation();
  if (persistConversationTimer) {
    clearTimeout(persistConversationTimer);
    persistConversationTimer = null;
  }
  activeConversationId = generateConversationId();
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
  chatLog.innerHTML = "";
  chatLog.dataset.empty = "true";
  insightsLog.textContent = "No topics logged yet. Start chatting to see trends.";
  hashtagCloud.innerHTML = "";
  metaTagSet.clear();
  updateMetaTags();
  syncMetaTagsFromKV(activeConversationId);
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
  html = html.replace(/\[(\d+)\]/g, (_, id) => `<a class="cite" data-cite-id="${id}" href="#" target="_blank" rel="noopener noreferrer">#${id}</a>`);
  html = highlightKeywords(html);
  if (needsMathWrapping(raw)) {
    html = wrapMath(html);
  }
  return html;
}

function looksLikeTableRow(line) {
  const normalized = normalizeTableLine(line);
  if (!normalized || !normalized.includes("|")) return false;
  const parts = normalized
    .replace(/^\s*\|/, "")
    .replace(/\|\s*$/, "")
    .split("|")
    .map(part => part.trim())
    .filter(Boolean);
  return parts.length >= 2;
}

function isMarkdownTableDivider(line) {
  if (!line) return false;
  const trimmed = normalizeTableLine(line).trim();
  if (!trimmed.includes("-")) return false;
  return /^\|?\s*:?-{2,}.*\|.*$/.test(trimmed);
}

function normalizeTableLine(line) {
  return (line || "").replace(/^\s*\\\|/, "|").replace(/\\\|/g, "|");
}

function splitMarkdownRow(line) {
  return normalizeTableLine(line)
    .replace(/^\s*\|/, "")
    .replace(/\|\s*$/, "")
    .split("|")
    .map(cell => cell.replace(/\\\|/g, "|").trim());
}

function parseMarkdownTableBlock(block) {
  if (!block || !block.includes("|")) return null;
  const lines = block
    .split(/\n/)
    .map(l => normalizeTableLine(l).trimEnd())
    .filter(Boolean);
  if (lines.length < 2) return null;

  let headerIdx = -1;
  for (let i = 0; i < lines.length - 1; i += 1) {
    if (looksLikeTableRow(lines[i]) && isMarkdownTableDivider(lines[i + 1])) {
      headerIdx = i;
      break;
    }
  }
  // Fallback: no explicit divider, but multiple table-like rows
  if (headerIdx === -1) {
    const rowLines = lines.filter(looksLikeTableRow);
    if (rowLines.length >= 2) {
      const caption = lines.length > rowLines.length ? lines.slice(0, lines.indexOf(rowLines[0])).join(" ") : "";
      const headerCells = splitMarkdownRow(rowLines[0]);
      const rows = rowLines.slice(1).map(splitMarkdownRow);
      if (!headerCells.length || !rows.length) return null;
      return { caption, headers: headerCells, rows };
    }
    return null;
  }

  const caption = lines.slice(0, headerIdx).filter(Boolean).join(" ");
  const headerCells = splitMarkdownRow(lines[headerIdx]);
  if (!headerCells.length) return null;

  const bodyLines = lines.slice(headerIdx + 2).filter(looksLikeTableRow);
  const rows = bodyLines.map(line => {
    const cells = splitMarkdownRow(line);
    const normalized = cells.slice(0, headerCells.length);
    while (normalized.length < headerCells.length) normalized.push("");
    return normalized;
  });

  return { caption, headers: headerCells, rows };
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
  const parsed = parseMarkdownTableBlock(block);
  if (!parsed) return null;

  const table = document.createElement("table");
  table.className = "owen-table";
  table.dataset.tone = pickTableTone(parsed.headers, parsed.rows);

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  parsed.headers.forEach(cellText => {
    const th = document.createElement("th");
    th.innerHTML = formatLine(cellText);
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  parsed.rows.forEach(row => {
    const tr = document.createElement("tr");
    row.forEach(cellText => {
      const td = document.createElement("td");
      td.innerHTML = formatLine(cellText || " ");
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  return {
    node: table,
    caption: parsed.caption,
    headingHint: parsed.caption || parsed.headers.join(" "),
  };
}

function isTableLikeBlock(block) {
  if (!block) return false;
  const lines = (block || "")
    .split(/\n/)
    .map(line => normalizeTableLine(line).trim())
    .filter(Boolean);
  if (lines.length < 2) return false;
  const pipeLines = lines.filter(line => line.includes("|"));
  if (pipeLines.length < 2) return false;
  const dividerCount = pipeLines.filter(isMarkdownTableDivider).length;
  const rowCount = pipeLines.filter(looksLikeTableRow).length;
  return rowCount >= 2 || (rowCount >= 1 && dividerCount >= 1);
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
    table.className = "owen-table";
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
    const headerCells = Array.from(table.querySelectorAll("thead th")).map(th => (th.textContent || "").trim());
    const palette = ["#60a5fa", "#34d399", "#fbbf24", "#f472b6", "#a78bfa", "#38bdf8"];
    headerCells.forEach((label, idx) => {
      const tab = document.createElement("span");
      tab.className = "owen-pretty-table__tab";
      tab.textContent = label || `Column ${idx + 1}`;
      tab.style.background = `linear-gradient(135deg, ${palette[idx % palette.length]}33, ${palette[(idx + 1) % palette.length]}55)`;
      tabs.appendChild(tab);
    });
    const wrap = document.createElement("div");
    wrap.className = "owen-pretty-table__wrap";
    wrap.appendChild(table);
    pretty.append(tabs, wrap);

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
      const h = document.createElement(level <= 2 ? "h4" : "h5");
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
      return expandCollapsedTables(
        seg
          .replace(/^\s*\\\|/gm, "|")
          .replace(/\\\|/g, "|"),
      );
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

function parseSectionsFromText(text) {
  const labelRe = new RegExp(`^(${SECTION_LABEL_WHITELIST_LOWER.join("|")}):\\s*(.*)$`, "i");
  const lines = text.split(/\r?\n/);
  const sections = [];
  let current = { label: "Response", lines: [] };
  lines.forEach(rawLine => {
    const line = rawLine || "";
    const match = labelRe.exec(line.trim());
    if (match) {
      if (current.lines.length) {
        sections.push({ ...current, text: current.lines.join("\n").trim() });
      }
      const nextLabel = match[1].replace(/\b\w/g, c => c.toUpperCase());
      const initial = match[2] ? [match[2]] : [];
      current = { label: nextLabel, lines: initial };
    } else {
      current.lines.push(line);
    }
  });
  if (current.lines.length || !sections.length) {
    sections.push({ ...current, text: current.lines.join("\n").trim() });
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

function resolveCitationLinks(root, citationMap) {
  if (!root) return;
  root.querySelectorAll("a.cite[data-cite-id]").forEach(link => {
    const id = link.getAttribute("data-cite-id");
    const url = id && citationMap.get(id);
    if (url) {
      link.href = url;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
    }
  });
}

function decorateAssistantResponse(text, citationMap = new Map()) {
  const sanitizedText = sanitizeAssistantTextForSections(text);
  enrichCitationMapFromText(text, citationMap);
  backfillCitationMapFromUrls(text, citationMap);
  const inlineUrls = collectUrlsFromText(text);
  const container = document.createElement("div");
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
    container.appendChild(p);
    return container;
  }
  const extractedReferences = [];
  const sections = parseSectionsFromText(trimmed);
  sections.forEach((sec, index) => {
    if (!sec || !sec.text) return;
    const section = document.createElement("div");
    section.className = "assistant-section";
    const heading = document.createElement("div");
    heading.className = "section-heading";
    heading.textContent = `${pickEmoji(sec.label)} ${sec.label}`;
    const body = document.createElement("div");
    body.className = "section-body";
    const blocks = sec.text
      .split(/\n{2,}/)
      .filter(Boolean)
      .filter(block => !isCitationsNoise(block));
    blocks.forEach(block => {
      if (isReferenceBlock(block)) return;
      const tableRender = buildMarkdownTable(block);
      if (tableRender) {
        if (tableRender.caption) {
          const caption = document.createElement("p");
          caption.className = "table-caption";
          caption.innerHTML = formatLine(tableRender.caption);
          body.appendChild(caption);
        }
        body.appendChild(tableRender.node);
      } else if (isTableLikeBlock(block)) {
        body.appendChild(renderTableFallback(block));
      } else {
        body.appendChild(renderMarkdownContent(block));
      }
    });
    resolveCitationLinks(body, citationMap);
    section.append(heading, body);
    container.appendChild(section);
    renderMathIfReady(section);
  });

  if (citationMap.size || extractedReferences.length || inlineUrls.length) {
    const refSection = document.createElement("div");
    refSection.className = "assistant-section";
    const refHeading = document.createElement("div");
    refHeading.className = "section-heading";
    refHeading.textContent = `${pickEmoji("reference")} References`;
    const refBody = document.createElement("div");
    refBody.className = "section-body";
    const entries = [];
    citationMap.forEach((url, id) => {
      if (isLikelyValidCitationUrl(url)) entries.push({ id, url });
    });
    let nextId = entries
      .map(entry => parseInt(entry.id, 10))
      .filter(n => !Number.isNaN(n))
      .reduce((max, n) => Math.max(max, n), entries.length ? 1 : 0) + 1;
    const seenUrls = new Set(entries.map(entry => entry.url));
    const addEntry = (url) => {
      const normalized = normalizeCitationUrl(url);
      if (!normalized || !isLikelyValidCitationUrl(normalized) || seenUrls.has(normalized)) return;
      const id = String(nextId++);
      entries.push({ id, url: normalized });
      seenUrls.add(normalized);
      citationMap.set(id, normalized);
    };
    extractedReferences.forEach(ref => addEntry(ref?.url));
    inlineUrls.forEach(addEntry);
    if (entries.length) {
      refBody.innerHTML = entries
        .map(entry => `<p id="ref-${escapeHTML(entry.id)}"><a class="cite" href="${escapeHTML(entry.url)}" target="_blank" rel="noopener noreferrer">#${entry.id}</a> ${escapeHTML(entry.url)}</p>`)
        .join("");
      refSection.append(refHeading, refBody);
      container.appendChild(refSection);
      renderMathIfReady(refSection);
    }
  }
  resolveCitationLinks(container, citationMap);
  return container;
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
  status.appendChild(statusText);
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

function startTypewriter(bubble, initialText = "", { onDone, completeOnSkip = false } = {}) {
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
  const textBlock = ensureBubbleTextNode(bubble, { reset: true });
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
  if (DEBUG_RENDER) {
    console.log("[RENDER_CHECK]", { msgId, answerLen: finalText.length, pipeLines });
    console.log("[RENDER_DEBUG]", { req: requestId, rawLen: finalText.length, table: looksLikeMarkdownTable(finalText) });
  }
  if (DEBUG_TABLE_RAW) {
    console.log("[RENDER_RAW_TEXT]", { msgId, requestId, text: finalText });
  }
  const renderPre = () => {
    const pre = document.createElement("pre");
    pre.className = "answer-pre";
    pre.textContent = finalText;
    return pre;
  };

  let node = null;
  try {
    node = decorateAssistantResponse(finalText, citationMap);
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
  if (!hashtagCloud) return [];
  const list = extractHashtags(promptText);
  if (!list.length) {
    hashtagCloud.innerHTML = "";
  } else {
    hashtagCloud.innerHTML = list
      .map(tag => `<span class="hashtag-chip">${escapeHTML(tag)}</span>`)
      .join("");
  }
  updateMetaTags(list);
  return list;
}

function renderMetaTagsBox(tags, counts) {
  if (!metaTagsBox) return;
  const list = Array.isArray(tags) ? tags.filter(Boolean) : [];
  if (!list.length) {
    metaTagsBox.dataset.empty = "true";
    metaTagsBox.textContent = "No question phrases yet.";
    return;
  }
  metaTagsBox.dataset.empty = "false";
  const countMap = counts instanceof Map ? counts : null;
  metaTagsBox.innerHTML = list
    .slice(0, 24)
    .map(tag => {
      const count = countMap ? countMap.get(tag) : null;
      const title = count ? ` title="${escapeHTML(String(count))} mentions"` : "";
      return `<span class="meta-tag"${title}>${escapeHTML(tag)}</span>`;
    })
    .join("");
}

function updateMetaTags(newTags) {
  if (!metaTagsBox || metaDataPanel) return;
  if (Array.isArray(newTags) && newTags.length) {
    newTags.forEach(tag => metaTagSet.add(tag));
  }
  if (!metaTagSet.size) {
    renderMetaTagsBox([]);
    return;
  }
  renderMetaTagsBox(Array.from(metaTagSet).slice(-24));
}

function rebuildMetaTagsFromConversation(messages) {
  metaTagSet.clear();
  if (Array.isArray(messages)) {
    messages.forEach(entry => {
      if (entry?.role === "user" && typeof entry.text === "string") {
        extractHashtags(entry.text).forEach(tag => metaTagSet.add(tag));
      }
    });
  }
  updateMetaTags();
}

async function syncMetaTagsFromKV(conversationId) {
  if (!conversationId) return;
  try {
    const res = await fetch(`/api/meta-tags?conversation_id=${encodeURIComponent(conversationId)}`);
    if (!res.ok) throw new Error("failed to fetch tags");
    const data = await res.json().catch(() => null);
    const tags = Array.isArray(data?.tags)
      ? data.tags.map(tag => (typeof tag === "string" ? tag : "")).filter(Boolean)
      : [];
    metaTagSet.clear();
    tags.forEach(tag => metaTagSet.add(tag));
    updateMetaTags();
  } catch (error) {
    console.warn("Could not sync meta tags", error);
    rebuildMetaTagsFromConversation(conversation);
  }
}

function normalizeMetaString(value) {
  if (typeof value !== "string") return "";
  const trimmed = value.trim();
  return trimmed || "";
}

function normalizeMetaStringCounts(rawStrings) {
  const counts = new Map();
  if (Array.isArray(rawStrings)) {
    rawStrings.forEach((entry) => {
      if (typeof entry === "string") {
        const text = normalizeMetaString(entry);
        if (text) counts.set(text, (counts.get(text) || 0) + 1);
        return;
      }
      if (entry && typeof entry === "object") {
        const text = normalizeMetaString(entry.text || entry.value || entry.tag || entry.label || "");
        const count = Number.isFinite(entry.count) ? Math.max(1, entry.count) : 1;
        if (text) counts.set(text, (counts.get(text) || 0) + count);
      }
    });
  }
  return Array.from(counts.entries())
    .map(([text, count]) => ({ text, count }))
    .sort((a, b) => (b.count - a.count) || a.text.localeCompare(b.text));
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

function setMetaStatusText(message, state) {
  if (!metaStatus) return;
  if (!message) {
    metaStatus.textContent = "";
    metaStatus.removeAttribute("data-state");
    return;
  }
  metaStatus.dataset.state = state || "info";
  metaStatus.textContent = message;
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
  setMetaStatusText("Waiting for selection.", "info");
  renderMetaTagsBox([]);
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
  const stringCounts = normalizeMetaStringCounts(payload.strings);
  const countMap = new Map(stringCounts.map(entry => [entry.text, entry.count]));
  renderMetaTagsBox(stringCounts.map(entry => entry.text), countMap);
  setMetaUpdated(payload.updated_at || "");
  setMetaStatusText("");
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
  setMetaStatusText("Loading...", "loading");
  setMetaErrorMessage("");
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
    if (!applyMetaAnalyticsPayload(payload, doc.docId, doc.title || doc.key || doc.docId)) {
      throw new Error("Meta data response did not match the selected lecture.");
    }
  } catch (error) {
    if (controller.signal.aborted) return;
    console.warn("Meta data fetch failed", error);
    renderMetaTagsBox([]);
    setMetaUpdated("");
    setMetaStatusText("");
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
    metaLectureSelect.innerHTML = `<option value="">Loading lectures…</option>`;
    return;
  }
  if (state === "error") {
    metaLectureSelect.innerHTML = `<option value="">Failed to load</option>`;
    resetMetaPanel();
    return;
  }
  metaLectureSelect.innerHTML = `<option value="">Select a lecture…</option>` + libraryListItems.map(item => {
    return `<option value="${escapeHTML(item.docId)}">${escapeHTML(item.title || item.key || item.docId)}</option>`;
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

function applyMetaUnlockState(isUnlocked) {
  metaDataUnlocked = isUnlocked;
  if (metaDataPanel) metaDataPanel.classList.toggle("is-locked", !isUnlocked);
  if (metaLockBtn) {
    metaLockBtn.dataset.state = isUnlocked ? "unlocked" : "locked";
    metaLockBtn.innerHTML = isUnlocked
      ? "<span aria-hidden=\"true\">🔓</span>"
      : "<span aria-hidden=\"true\">🔒</span>";
    metaLockBtn.setAttribute("aria-label", isUnlocked ? "Lock meta data" : "Unlock meta data");
  }
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
  } else {
    refreshMetaDataPanel();
  }
}

const formatTime = () => new Intl.DateTimeFormat(undefined, {
  hour: "2-digit",
  minute: "2-digit",
}).format(new Date());

function setStatus(el, state, message) {
  if (!message) {
    el.removeAttribute("data-state");
    el.textContent = "";
    return;
  }
  el.dataset.state = state || "info";
  el.textContent = message;
}

function updateSendButtonState() {
  const hasName = nameInput.value.trim().length > 0;
  const hasFile = Boolean(fileInput.files && fileInput.files.length);
  sendButton.disabled = !(hasName && hasFile);
}

function updateAttachmentIndicator() {
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
        <button type="button" data-remove="oneshot" aria-label="Remove attachment">✕</button>
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
    return;
  }
  if (!attachments.length) {
    attachStatus.dataset.state = "empty";
    attachStatus.textContent = "No files attached";
    attachStatus.removeAttribute("title");
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
        <button type="button" data-remove="${idx}" aria-label="Remove attachment">✕</button>
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
  } = options;
  chatLog.dataset.empty = "false";
  const bubble = document.createElement("div");
  bubble.className = `bubble ${role}`;
  const resolvedMsgId = msgId || generateMessageId();
  bubble.dataset.msgId = resolvedMsgId;
  if (docId) bubble.dataset.docId = docId;
  if (requestId) bubble.dataset.requestId = requestId;
  const textBlock = document.createElement("div");
  textBlock.className = "bubble-text";
  if (role === "assistant") {
    try {
      textBlock.replaceChildren(decorateAssistantResponse(text, new Map()));
    } catch (err) {
      console.error("[RENDER_ERROR_INIT]", { err, answerLen: (text || "").length, msgId: resolvedMsgId });
      const pre = document.createElement("pre");
      pre.style.whiteSpace = "pre-wrap";
      pre.textContent = text || "";
      textBlock.replaceChildren(pre);
    }
    enhanceAutoTables(textBlock);
    enhanceTablesIn(textBlock);
    renderReferenceChips(bubble, references);
    if (evidence && evidence.length) {
      renderEvidenceToggle(bubble, evidence);
    }
  } else {
    textBlock.textContent = text;
  }
  bubble.appendChild(textBlock);

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
  const isLight = mode === "light";
  document.body.classList.toggle("theme-light", isLight);
  document.documentElement.setAttribute("data-theme", isLight ? "light" : "dark");
  themeToggle.setAttribute("aria-pressed", String(isLight));
  themeToggle.innerHTML = isLight ? THEME_TOGGLE_MOON_ICON : THEME_TOGGLE_LIGHT_ICON;
}

const savedTheme = localStorage.getItem(THEME_KEY);
applyTheme(savedTheme === "light" ? "light" : "dark");

const UPLOAD_COLLAPSE_KEY = "owen.uploadPanelCollapsed";
const UPLOAD_UNLOCK_KEY = "owen.uploadUnlocked";
const META_UNLOCK_KEY = "owen.metaDataUnlocked";
const HISTORY_COLLAPSE_KEY = "owen.chatHistoryCollapsed";
if (uploadPanel && uploadToggle && uploadPanelBody) {
  const collapsed = localStorage.getItem(UPLOAD_COLLAPSE_KEY) === "true";
  const unlocked = localStorage.getItem(UPLOAD_UNLOCK_KEY) === "true";
  if (collapsed) uploadPanel.classList.add("is-collapsed");
  if (!unlocked) uploadPanel.classList.add("is-locked");
  const applyUnlockState = (isUnlocked) => {
    uploadPanel.classList.toggle("is-locked", !isUnlocked);
    if (uploadLockBtn) uploadLockBtn.hidden = !isUnlocked;
    if (!isUnlocked) uploadPanel.classList.add("is-collapsed");
    try {
      localStorage.setItem(UPLOAD_UNLOCK_KEY, String(isUnlocked));
    } catch {}
  };
  applyUnlockState(unlocked);
  if (uploadToggle) {
    uploadToggle.addEventListener("click", () => {
      if (uploadPanel.classList.contains("is-locked")) {
        uploadUnlockInput?.focus();
        return;
      }
      const next = !uploadPanel.classList.contains("is-collapsed");
      uploadPanel.classList.toggle("is-collapsed", next);
      try {
        localStorage.setItem(UPLOAD_COLLAPSE_KEY, String(next));
      } catch {}
    });
  }
  if (uploadUnlockInput) {
    const tryUnlock = () => {
      if (uploadUnlockInput.value === "1234") {
        applyUnlockState(true);
        uploadUnlockInput.value = "";
      }
    };
    uploadUnlockInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        tryUnlock();
      }
    });
    uploadUnlockInput.addEventListener("blur", () => {
      if (uploadUnlockInput.value === "1234") {
        applyUnlockState(true);
        uploadUnlockInput.value = "";
      }
    });
  }
  if (uploadLockBtn) {
    uploadLockBtn.addEventListener("click", () => {
      applyUnlockState(false);
      try {
        localStorage.setItem(UPLOAD_COLLAPSE_KEY, "true");
      } catch {}
    });
  }
}

if (metaDataPanel && metaUnlockInput && metaLockBtn) {
  const unlocked = localStorage.getItem(META_UNLOCK_KEY) === "true";
  applyMetaUnlockState(unlocked);
  const tryUnlock = () => {
    if (metaUnlockInput.value === "1234") {
      applyMetaUnlockState(true);
      metaUnlockInput.value = "";
    }
  };
  metaUnlockInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      tryUnlock();
    }
  });
  metaUnlockInput.addEventListener("blur", tryUnlock);
  metaLockBtn.addEventListener("click", () => {
    if (metaDataPanel.classList.contains("is-locked")) {
      tryUnlock();
      return;
    }
    applyMetaUnlockState(false);
    metaUnlockInput.value = "";
  });
}

if (metaLectureSelect) {
  metaLectureSelect.addEventListener("change", () => refreshMetaDataPanel());
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

if (fullDocToggle) {
  fullDocToggle.checked = true;
  fullDocToggle.addEventListener("change", () => {
    fullDocumentMode = Boolean(fullDocToggle.checked);
  });
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

themeToggle.addEventListener("click", () => {
  const next = document.body.classList.contains("theme-light") ? "dark" : "light";
  applyTheme(next);
  localStorage.setItem(THEME_KEY, next);
});

runRenderFallbackSelfTest();
runEmptyFinalRegressionTest();

fileButton.addEventListener("click", () => fileInput.click());
chatAttach.addEventListener("click", () => {
  chatFileInput.value = "";
  chatFileInput.click();
});

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

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  fileNameDisplay.textContent = file ? file.name : "No file selected";
  updateSendButtonState();
});

nameInput.addEventListener("input", updateSendButtonState);

async function uploadFile() {
  const subject = nameInput.value.trim();
  const file = fileInput.files?.[0] ?? null;
  if (!subject) {
    setStatus(uploadStatus, "error", "Please enter a name before uploading.");
    nameInput.focus();
    return;
  }
  if (!file) {
    setStatus(uploadStatus, "error", "Choose a file to upload.");
    fileButton.focus();
    return;
  }

  sendButton.disabled = true;
  sendButton.textContent = "Sending...";
  fileButton.disabled = true;
  setStatus(uploadStatus, "pending", "Uploading to R2 + mirroring to OpenAI Files...");

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
      let statusMsg = `${displayName} saved to ${payload.bucket}/${payload.key}.`;
      if (payload.vector_store_id) {
        statusMsg += ragReady
          ? " Retrieval index ready for grounded answers."
          : ' Indexing in progress; answers may say "I don\'t know" until processing finishes.';
      } else {
        statusMsg += " Attached, but automatic indexing failed — retry if you need RAG answers.";
      }
      if (payload.ocr_status === "ready") {
        statusMsg += " OCR transcript ready for exact quotations.";
      } else if (payload.ocr_status === "error") {
        statusMsg += ` OCR failed: ${payload.ocr_warning || "Unknown error"}.`;
        statusState = "warning";
      } else if (payload.ocr_status === "empty") {
        statusMsg += " OCR produced no text; the PDF might be blank.";
      }
      if (payload.vector_warning) {
        statusMsg += ` (${payload.vector_warning})`;
        statusState = "warning";
      }
      setStatus(uploadStatus, statusState, statusMsg);
    } else {
      setStatus(
        uploadStatus,
        "error",
        "Uploaded to R2 but failed to mirror to OpenAI. Attach via the + button to retry.",
      );
    }
    fileInput.value = "";
    fileNameDisplay.textContent = "No file selected";
    updateSendButtonState();
  } catch (error) {
    console.error(error);
    setStatus(uploadStatus, "error", error instanceof Error ? error.message : "Unexpected upload error.");
  } finally {
    sendButton.textContent = "Send";
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

sendButton.addEventListener("click", uploadFile);

async function uploadChatAttachment(file) {
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

async function ingestPdfToWorker(file, arrayBuffer, fileHash, pageCap = OCR_PAGE_LIMIT) {
  const form = new FormData();
  form.append("file", file);
  form.append("fileHash", fileHash);
  form.append("maxPages", String(pageCap));
  const resp = await fetch("/api/pdf-ingest", { method: "POST", body: form });
  const payload = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const msg = payload?.error || payload?.details || "PDF ingest failed.";
    throw new Error(msg);
  }
  return payload;
}

async function askDocWithRetrieval(question, extractedKey, fileHash) {
  const resp = await fetch("/api/ask-doc", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ question, extractedKey, fileHash }),
  });
  const payload = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const msg = payload?.error || payload?.details || "Document Q&A failed.";
    throw new Error(msg);
  }
  return payload?.answer || "(no answer returned)";
}

async function ocrPageImage({ fileHash, fileId, pageIndex, blob, filename }) {
  let attempt = 0;
  let lastError = "";
  while (attempt < OCR_MAX_RETRIES) {
    const fd = new FormData();
    if (fileHash) fd.append("fileHash", fileHash);
    if (fileId) fd.append("fileId", fileId);
    fd.append("pageIndex", String(pageIndex));
    fd.append("image", blob, filename || `page-${pageIndex + 1}.jpg`);
    const resp = await fetch("/api/ocr-page", { method: "POST", body: fd });
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

async function finalizeOcrText({ fileHash, fileId, pages, filename, totalPages }) {
  const resp = await fetch("/api/ocr-finalize", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ fileHash, fileId, pages, filename, totalPages }),
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

async function runPdfOcrFlow({ file, prompt, pageCap = OCR_PAGE_LIMIT, thinkingBubble, startPage = 1 }) {
  const wantsAllPages = wantsEntireDocument(prompt) || fullDocumentMode;
  const requestedCap = wantsAllPages ? Math.max(pageCap, 9999) : pageCap;
  setThinkingStatus(thinkingBubble, "Hashing PDF…");
  const arrayBuffer = await file.arrayBuffer();
  const fileHash = await computeSha256(arrayBuffer);
  setThinkingStatus(thinkingBubble, "Checking PDF…");
  const ingest = await ingestPdfToWorker(file, arrayBuffer, fileHash, requestedCap);
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
    const answer = await askDocWithRetrieval(prompt, ingest.extractedKey, fileHash);
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
  const pdfDoc = await pdfjsLib.getDocument({
    data: arrayBuffer,
    disableWorker: true,
  }).promise;
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
      });
      state.extractedKey = finalize.extractedKey;
      state.partialPages = ordered;
      state.answer = await askDocWithRetrieval(prompt, finalize.extractedKey, fileHash);
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
      throw new Error("Processing cancelled.");
    }
    if (thinkingBubble && thinkingBubble.isConnected) {
      setThinkingStatus(thinkingBubble, `OCR page ${state.allPages.length + 1} / ${pageNumbers.length}`);
    }
    const render = await renderPdfPageToBlob(pdfjsLib, pdfDoc, pageNumber);
    const pageResult = await ocrPageImage({
      fileHash,
      fileId: ingest?.fileId || null,
      pageIndex: pageNumber - 1,
      blob: render.blob,
      filename: file.name,
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
      console.error("[OCR] background failure", err);
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
  const continueRange = !wantsAllPages && totalPages > end
    ? { start: end + 1, end: Math.min(totalPages, end + OCR_PAGE_LIMIT), totalPages }
    : null;
  const warning = `Processed pages ${start}-${end}.`;
  return { answer: state.answer, warning, continueRange, extractedKey: state.extractedKey, fileHash };
}

async function continuePdfRange({ start, end, prompt, thinkingBubble }) {
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
  const pdfjsLib = await loadPdfJsLib();
  let pdfDoc = null;
  try {
    pdfDoc = await pdfjsLib.getDocument({
      data: activePdfSession.arrayBuffer,
      disableWorker: true,
    }).promise;
    const totalPages = activePdfSession.pageCount || pdfDoc?.numPages || end;
    const safeEnd = Math.min(end, totalPages);
    const pageNumbers = [];
    for (let i = start; i <= safeEnd; i += 1) pageNumbers.push(i);
    setThinkingStatus(thinkingBubble, `Processing pages ${start}-${safeEnd} of ${totalPages}`);
    const results = [];
    await promisePool(pageNumbers, OCR_CONCURRENCY, async (pageNumber) => {
      if (thinkingBubble && thinkingBubble.isConnected) {
        setThinkingStatus(thinkingBubble, `OCR page ${results.length + 1} / ${pageNumbers.length}`);
      }
      const render = await renderPdfPageToBlob(pdfjsLib, pdfDoc, pageNumber);
      const pageResult = await ocrPageImage({
        fileHash: activePdfSession.fileHash,
        fileId: null,
        pageIndex: pageNumber - 1,
        blob: render.blob,
        filename: activePdfSession.filename,
      });
      results.push(pageResult);
    });
    results.sort((a, b) => a.pageIndex - b.pageIndex);
    const finalize = await finalizeOcrText({
      fileHash: activePdfSession.fileHash,
      pages: results,
      filename: activePdfSession.filename,
      totalPages,
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
    );
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
    try {
      const result = await continuePdfRange({
        start: range.start,
        end: range.end,
        prompt,
        thinkingBubble: thinking,
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
      console.error(err);
      if (thinking && thinking.isConnected) thinking.remove();
      const errMsg = err instanceof Error ? err.message : "Continue OCR failed.";
      const errBubble = appendChatMessage("assistant", errMsg, { track: false });
      errBubble.classList.add("error");
    }
  });
  cancelBtn.addEventListener("click", () => {
    cancelBtn.disabled = true;
    activePdfSession = activePdfSession ? { ...activePdfSession, cancelled: true } : null;
    persistOcrSession(activePdfSession);
    bubble.remove();
  });
}

async function askFileWithFallback({ prompt, file, thinkingBubble }) {
  if (isPdfFile(file)) {
    return runPdfOcrFlow({ file, prompt, pageCap: OCR_PAGE_LIMIT, thinkingBubble });
  }
  const fd = new FormData();
  fd.append("message", prompt);
  fd.append("file", file);
  const response = await fetch("/api/ask-file", { method: "POST", body: fd });
  const payload = await response.json().catch(() => ({}));
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
        opt.textContent = doc.title || doc.key || doc.docId;
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
        opt.textContent = doc.title || doc.key || doc.docId;
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
    selectedDocTitle: doc?.title || doc?.key || "",
  };
  schedulePersistConversation(200);
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
    li.innerHTML = `
      <div class="library-row-top">
        <h4>${escapeHTML(rec.title || rec.key || rec.docId)}</h4>
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
        setStatus(libraryStatus, "success", `Using cached lecture: ${rec.title}`);
        if (chatInput.value.trim()) {
          try {
            await askLibraryDocQuestion({ ...rec, status: "ready" }, chatInput.value.trim());
          } catch {
            // handled upstream
          }
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

async function askLibraryDocQuestion(doc, prompt) {
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
  const userBubble = appendChatMessage("user", question);
  renderMathIfReady(userBubble);
  const thinking = showThinkingBubble();
  setBubbleStatus(thinking, "O.W.E.N. Is Thinking");
  try {
    const resp = await fetch("/api/library/ask", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ docId: doc.docId, question, requestId }),
    });
    const payload = await resp.json().catch(() => ({}));
    const normalized = normalizeLibraryAskPayload(payload, allowExcerpts);
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
        (normalized.error && normalized.error.message) ||
        payload?.error ||
        payload?.details ||
        `Library Q&A failed (status ${resp.status}).`;
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
      docTitle: doc.title || doc.key || doc.docId,
      requestId: normalized.debug?.requestId || payload?.requestId || requestId,
      references: !isTable ? normalized.references : undefined,
      evidence: !isTable && allowExcerpts ? normalized.evidence : undefined,
    });
    setStatus(libraryStatus, "success", `Answered using ${doc.title || doc.docId}.`);
    setLibrarySelection({ ...doc, status: "ready" });
    chatInput.value = "";
    return payload;
  } catch (error) {
    console.error(error);
    if (thinking && thinking.isConnected) thinking.remove();
    const bubble = appendChatMessage("assistant", error instanceof Error ? error.message : "Unable to answer from library.", { track: false });
    bubble.classList.add("error");
    setStatus(libraryStatus, "error", error instanceof Error ? error.message : "Library Q&A failed.");
    throw error;
  } finally {
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

function normalizeReferenceEntries(references, { docId } = {}) {
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
  const { items, fallbackText } = normalizeReferenceEntries(references, { docId });
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
    setLibrarySelection(readyDoc);
    setStatus(libraryStatus, "success", "Cached text ready.");
    await runLibrarySearch();
    await loadLibraryList();
    if (chatInput.value.trim()) {
      await askLibraryDocQuestion(readyDoc, chatInput.value.trim());
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
    await runLibrarySearch();
    await loadLibraryList();
  } catch (error) {
    console.error(error);
    setStatus(libraryStatus, "error", error instanceof Error ? error.message : "Batch ingest failed.");
  }
}

async function loadLibraryList() {
  if (!librarySelect || !libraryStatusBadge) return;
  librarySelect.innerHTML = `<option value="">Loading lectures…</option>`;
  updateMetaLectureOptions("loading");
  libraryStatusBadge.dataset.state = "info";
  libraryStatusBadge.textContent = "Loading…";
  try {
    const resp = await fetch("/api/library/list?prefix=library/");
    const payload = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(payload?.error || "Failed to load library list.");
    libraryListItems = Array.isArray(payload?.items) ? payload.items : [];
    libraryListItems.sort((a, b) => (a.title || "").localeCompare(b.title || ""));
    librarySelect.innerHTML = `<option value="">Select a lecture…</option>` + libraryListItems.map(item => {
      return `<option value="${escapeHTML(item.docId)}">${escapeHTML(item.title || item.key || item.docId)}</option>`;
    }).join("");
    updateMetaLectureOptions();
    libraryStatusBadge.dataset.state = "info";
    libraryStatusBadge.textContent = "Pick a lecture";
  } catch (err) {
    console.error(err);
    librarySelect.innerHTML = `<option value="">Failed to load</option>`;
    libraryStatusBadge.dataset.state = "error";
    libraryStatusBadge.textContent = err instanceof Error ? err.message : "Load failed";
    updateMetaLectureOptions("error");
  }
}

function updateLibraryUiForSelection(doc) {
  if (libraryStatusBadge) {
    if (!doc) {
      libraryStatusBadge.dataset.state = "info";
      libraryStatusBadge.textContent = "No selection";
    } else {
      const state = doc.status === "ready" ? "success" : "warning";
      libraryStatusBadge.dataset.state = state;
      libraryStatusBadge.textContent = doc.status === "ready" ? "READY" : "MISSING";
    }
  }
  if (libraryActionBtn) {
    libraryActionBtn.disabled = !doc;
    libraryActionBtn.textContent = doc?.status === "ready" ? "Use cached lecture" : "Process once";
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
  if (activeLibraryDoc.status === "ready") {
    setStatus(libraryStatus, "success", `Using cached lecture: ${activeLibraryDoc.title || activeLibraryDoc.key}`);
    return;
  }
  await triggerLibraryIngest(activeLibraryDoc);
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
  if (handleLocalCommands(prompt)) return;
  if (activeLibraryDoc && attachments.length === 0 && !oneShotFile) {
    try {
      await askLibraryDocQuestion(activeLibraryDoc, prompt);
    } finally {
      chatStreaming = false;
      chatSend.disabled = false;
    }
    return;
  }
  if (chatStreaming) return;
  chatStreaming = true;
  chatSend.disabled = true;
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
    appendChatMessage("user", prompt, { attachments: attachmentsForBubble });
    chatInput.value = "";
    insightsLog.textContent = `Last ask (${model}): ${prompt}`;
    const thinkingBubble = showThinkingBubble();
    try {
      const result = await askFileWithFallback({ prompt, file: oneShotFile, thinkingBubble });
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
      console.error(error);
      if (thinkingBubble && thinkingBubble.isConnected) thinkingBubble.remove();
      const bubble = appendChatMessage("assistant", error instanceof Error ? error.message : "Something went wrong.");
      bubble.classList.add("error");
    } finally {
      chatStreaming = false;
      chatSend.disabled = false;
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
  appendChatMessage("user", prompt, { attachments: attachmentsSnapshot });
  chatInput.value = "";
  insightsLog.textContent = `Last ask (${model}): ${prompt}`;
  const thinkingBubble = showThinkingBubble();
  setBubbleStatus(thinkingBubble, "Queued…");
  let live = thinkingBubble;
  let hasContent = false;
  let acc = "";
  const citationEvents = [];
  const msgId = thinkingBubble?.dataset?.msgId || generateMessageId();
  if (thinkingBubble) thinkingBubble.dataset.msgId = msgId;
  messageState.set(msgId, { streamText: "", lastAnswer: "" });

  try {
    setBubbleStatus(live, "O.W.E.N. Is Thinking");
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (isImageModel(model)) {
      const imagePayload = await response.json().catch(() => ({}));
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
    if (!response.body) throw new Error("Streaming response missing body");
    const reader = response.body.getReader();
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
              ensureBubbleTextNode(live, { reset: true });
              startTypewriter(live, acc);
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
      const { done, value } = await reader.read();
      if (value) buffer += decoder.decode(value, { stream: true });
      flushBuffer();
      if (done) break;
    }
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
      renderFinalAssistant(live, normalizedText, citeMap);
      logMessage("assistant", normalizedText, {
        model,
        id: msgId,
        docId: live?.dataset?.docId,
        requestId: live?.dataset?.requestId,
        docTitle: activeConversationMeta.selectedDocTitle,
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
    console.error(error);
    if (live && live.isConnected) live.remove();
    const bubble = appendChatMessage("assistant", error instanceof Error ? error.message : "Something went wrong while chatting.");
    bubble.classList.add("error");
  } finally {
    chatStreaming = false;
    chatSend.disabled = false;
  }
}

chatForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const prompt = chatInput.value.trim();
  if (!prompt) {
    chatInput.focus();
    return;
  }
  runChat(prompt);
});

chatInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    const prompt = chatInput.value.trim();
    if (prompt) runChat(prompt);
  }
});

chatInput.addEventListener("paste", handleChatPaste);

chatLog.addEventListener("click", (event) => {
  if ((event.target instanceof Element) && event.target.closest(".bubble.assistant")) {
    skipAllTypewriters();
  }
});

chatLog.addEventListener("wheel", () => skipAllTypewriters(), { passive: true });
window.addEventListener("keydown", (event) => {
  if (event.key === " " || event.key === "Enter") {
    skipAllTypewriters();
  }
});

retrieveBtn.addEventListener("click", async () => {
  const key = retrieveInput.value.trim();
  if (!key) {
    setStatus(retrieveStatus, "error", "Enter a filename or storage key first.");
    retrieveInput.focus();
    return;
  }
  await retrieveFileFromR2({ key, bucket: "", autoDownload: true, statusEl: retrieveStatus, emitChat: true });
});

if (librarySearchBtn) librarySearchBtn.addEventListener("click", () => runLibrarySearch());
if (librarySearchInput) librarySearchInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    runLibrarySearch();
  }
});
if (libraryRefreshBtn) libraryRefreshBtn.addEventListener("click", () => loadLibraryList());
if (librarySelect) librarySelect.addEventListener("change", onLibrarySelectChange);
if (libraryActionBtn) libraryActionBtn.addEventListener("click", onLibraryAction);
if (libraryClearBtn) libraryClearBtn.addEventListener("click", () => {
  setLibrarySelection(null);
  if (librarySelect) librarySelect.value = "";
});
if (libraryIndexBtn) libraryIndexBtn.addEventListener("click", runLibraryBatchIndex);
if (libraryIngestBtn) libraryIngestBtn.addEventListener("click", (event) => {
  const mode = event.shiftKey || event.metaKey ? "full" : "embedded_only";
  runLibraryBatchIngest(mode);
});

window.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    clearConversation();
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

updateSendButtonState();
updateAttachmentIndicator();
modelButtons.forEach(btn => {
  btn.addEventListener("click", () => {
    // Model selection disabled; keep currentModel fixed.
  });
});

renderConversationList();
const lastSavedId = loadActiveConversationId();
if (lastSavedId) {
  loadConversation(lastSavedId);
} else if (storedConversations.length) {
  loadConversation(storedConversations[0].id);
}
newConvoBtn.addEventListener("click", () => startNewConversation());
syncMetaTagsFromKV(activeConversationId);
setLibrarySelection(null);
loadLibraryList().catch(() => {});
runLibrarySearch().catch(() => {});
