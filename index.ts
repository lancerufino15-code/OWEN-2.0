/**
 * Legacy/alternate Worker entrypoint retained alongside the `src/` version.
 *
 * Used by: not referenced in current Wrangler configs (see `wrangler.jsonc`),
 * but kept for historical compatibility or local experiments.
 *
 * Key exports:
 * - Default `fetch` handler for Worker runtime.
 * - `ChatMessage` type for request shaping.
 *
 * Assumptions:
 * - Cloudflare Workers runtime with R2 + KV bindings configured via Env.
 */
import { AGENTS } from "./agents";
import type { OwenAgent } from "./agents";
import type { Env } from "./types";
import * as pdfjs from "./pdfjs-dist-legacy-build-pdf";
import { normalizePages, normalizePlainText, type PageText } from "./pdf/normalize";
import {
  buildExtractedKeyForHash,
  loadCachedExtraction,
  readManifest,
  writeManifest,
  type PdfManifest,
} from "./pdf/cache";
import { chunkText, rankChunks, type RetrievalChunk } from "./pdf/retrieval";
import {
  LIBRARY_INDEX_KEY,
  LIBRARY_QUEUE_PREFIX,
  buildExtractedPath,
  buildIndexKeyForDoc,
  buildManifestPath,
  computeDocId,
  isPdfKey,
  type LibraryIndexRecord,
  normalizePreview,
  readIndex as readLibraryIndex,
  scoreRecords as scoreLibraryRecords,
  titleFromKey,
  tokensFromTitle,
  writeIndex as writeLibraryIndex,
} from "./library";
import {
  MAX_TAGS,
  aggregateLectureAnalytics,
  buildAnalyticsEvent,
  loadTopQuestionsForLecture,
  writeAnalyticsEvent,
} from "./analytics";
import { filterMetadata } from "./analytics_metadata_filter";

// Cloudflare Workers cannot spawn PDF.js workers; force single-threaded parsing.
if (pdfjs?.GlobalWorkerOptions) {
  pdfjs.GlobalWorkerOptions.workerSrc = undefined;
}
const PDFJS_AVAILABLE = Boolean(pdfjs?.getDocument) && !(pdfjs as any).__isStub;

type ConversationTags = {
  conversation_id: string;
  tags: string[];
  updated_at: string;
};

type LectureAnalytics = {
  docId: string;
  tags: Array<{ tag: string; count: number }>;
  updated_at: string;
  entities?: Array<{ entity: string; count: number }>;
  questions?: Array<{ question: string; count: number }>;
  lastUpdated?: string;
};

type ChatRole = "system" | "user" | "assistant";

type ChatContentPart = { type: "text"; text: string };

/**
 * Minimal chat message shape used for OpenAI-style chat payloads.
 */
export type ChatMessage = { role: ChatRole; content: string | ChatContentPart[] };

type InputImageDetail = "low" | "high" | "auto";
type ResponsesInputContent =
  | { type: "input_text"; text: string }
  | { type: "output_text"; text: string }
  | { type: "input_image"; image_url?: string; file_id?: string; detail?: InputImageDetail }
  | { type: "input_file"; file_id: string };

type ResponsesInputMessage = { role: ChatRole; content: ResponsesInputContent[] };

interface FileReference {
  bucket: string;
  key: string;
  textKey?: string;
  displayName?: string;
  fileId?: string;
  visionFileId?: string;
}

type FileContextRecord = {
  displayName: string;
  source: "ocr" | "original";
  text: string;
  bucket: string;
  resolvedBucket: string;
  originalKey: string;
  resolvedKey: string;
  textKey?: string;
};

interface ChatRequestBody {
  messages: ChatMessage[];
  agentId?: string;
  model?: string;
  files?: FileReference[];
  attachments?: FileReference[];
  fileRefs?: FileReference[];
  conversation_id?: string;
  meta_tags?: string[];
}

interface GenerateFileRequest {
  messages: ChatMessage[];
  agentId?: string;
  desiredFileType: string;
}

async function saveConversationTags(env: Env, conversation_id: string, rawTags: unknown): Promise<void> {
  if (!conversation_id || !env.DOCS_KV) return;

  const tags: string[] = Array.isArray(rawTags)
    ? [...new Set(
        rawTags
          .map((t) => String(t).trim())
          .filter(Boolean)
          .map((t) => (t.startsWith("#") ? t.toLowerCase() : `#${t.toLowerCase()}`))
      )]
    : [];

  if (tags.length === 0) return;

  const kvKey = `conv:${conversation_id}:tags`;

  const existing = await env.DOCS_KV.get(kvKey, { type: "json" }) as ConversationTags | null;

  const mergedTags = existing
    ? [...new Set([...(existing.tags || []), ...tags])]
    : tags;

  const record: ConversationTags = {
    conversation_id,
    tags: mergedTags,
    updated_at: new Date().toISOString(),
  };

  await env.DOCS_KV.put(kvKey, JSON.stringify(record));
}

async function loadConversationTags(env: Env, conversation_id: string): Promise<ConversationTags> {
  if (!env.DOCS_KV) {
    return { conversation_id, tags: [], updated_at: new Date().toISOString() };
  }
  const kvKey = `conv:${conversation_id}:tags`;
  const stored = await env.DOCS_KV.get(kvKey, { type: "json" }) as ConversationTags | null;

  if (!stored) {
    return { conversation_id, tags: [], updated_at: new Date().toISOString() };
  }

  return {
    conversation_id: stored.conversation_id || conversation_id,
    tags: Array.isArray(stored.tags) ? stored.tags : [],
    updated_at: stored.updated_at || new Date().toISOString(),
  };
}

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, content-type",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
};
const NO_STORE_HEADERS = {
  "cache-control": "no-store, max-age=0",
};
const FACULTY_PASSCODE_PLACEHOLDER = "__FACULTY_PASSCODE__";

function resolveFacultyPasscode(env: Env): string {
  const candidates = [
    env.FACULTY_PASSCODE,
    env.NEXT_PUBLIC_FACULTY_PASSCODE,
    env.VITE_FACULTY_PASSCODE,
  ];
  for (const entry of candidates) {
    if (typeof entry === "string" && entry.trim()) {
      return entry.trim();
    }
  }
  return "";
}

async function injectFacultyPasscode(resp: Response, env: Env): Promise<Response> {
  const contentType = resp.headers.get("content-type") || "";
  if (!contentType.includes("text/html")) return resp;
  const text = await resp.text();
  const passcode = resolveFacultyPasscode(env);
  const updated = text.includes(FACULTY_PASSCODE_PLACEHOLDER)
    ? text.split(FACULTY_PASSCODE_PLACEHOLDER).join(passcode)
    : text;
  const headers = new Headers(resp.headers);
  headers.set("content-type", "text/html; charset=utf-8");
  headers.delete("content-length");
  headers.delete("etag");
  return new Response(updated, { status: resp.status, statusText: resp.statusText, headers });
}

const ALLOWED_MODELS = [
  "gpt-5-mini",
  "gpt-4o",
  "gpt-4-turbo",
  "gpt-4.1",
  "gpt-4.1-mini",
  "gpt-5",
  "gpt-image-1",
  "gpt-image-1-mini",
  "dall-e-3",
  "dall-e-2",
] as const;
type AllowedModel = (typeof ALLOWED_MODELS)[number];

const BUCKET_BINDINGS = {
  "owen-bucket": "OWEN_BUCKET",
  "owen-ingest": "OWEN_INGEST",
  "owen-notes": "OWEN_NOTES",
  "owen-test": "OWEN_TEST",
  "owen-uploads": "OWEN_UPLOADS",
  "own-ingest": "OWN_INGEST",
} as const;

const DEFAULT_BUCKET = "owen-uploads" as const;
const CANONICAL_CACHE_BUCKET_NAME = "owen-ingest" as const;
const EXTRACTION_BUCKET = CANONICAL_CACHE_BUCKET_NAME;
const LIBRARY_BUCKET = CANONICAL_CACHE_BUCKET_NAME;
const enc = new TextEncoder();
const dec = new TextDecoder();
const ANSWER_MAX_OUTPUT_TOKENS = 5000;
const ANSWER_MAX_CONTINUATIONS = 6;
const CONTINUATION_TAIL_CHARS = 2000;
const LIBRARY_NOTES_CONTEXT_CHARS = 12_000;
const LIBRARY_NOTES_MAX_OUTPUT_TOKENS = 1000;
const LIBRARY_TOTAL_CONTINUATIONS = 24;
const TRUNCATION_NOTICE = "(Response truncated due to server time limits -- click 'Continue' to fetch the rest.)";
const VECTOR_POLL_INTERVAL_MS = 1500;
const VECTOR_POLL_TIMEOUT_MS = 45_000;
const VECTOR_STORE_CACHE_KEY = "owen.vector_store_id";
let inMemoryVectorStoreId: string | null = null;
const FILE_CONTEXT_CHAR_LIMIT = 60_000;
const MAX_OCR_TEXT_LENGTH = 900_000;
const MAX_STORED_TEXT_LENGTH = 900_000;
const DEFAULT_MAX_OCR_PAGES = 200;
const CLIENT_OCR_PAGE_CAP = 15;
const MIN_EMBEDDED_PDF_CHARS = 800;
const PDF_SAMPLE_PAGE_TARGET = 5;
const ONE_SHOT_PREVIEW_LIMIT = 1200;
const MAX_EXTRACT_CHARS = 120_000;
const MAX_EXTRACT_PAGES = 200;
const OCR_MAX_OUTPUT_TOKENS = 3000;
const OCR_PAGE_OUTPUT_TOKENS = 3000;
const RETRIEVAL_CHUNK_SIZE = 2200;
const RETRIEVAL_CHUNK_OVERLAP = 200;
const RETRIEVAL_TOP_K = 6;
const LIBRARY_BROAD_TOP_K = 32;
const LIBRARY_BROAD_MIN_CHUNKS = 12;
const LIBRARY_BROAD_MIN_CHARS = 12_000;
const LIBRARY_BROAD_MAX_CONTEXT_CHARS = 36_000;
const DEBUG_FORMATTING = false;
const RETRY_DELAYS_MS = [0, 500, 1500];
const DEFAULT_OCR_MAX_PAGES = 15;
const LIBRARY_SEARCH_LIMIT = 12;
const LIBRARY_BATCH_PAGE_LIMIT = 5;
const LIBRARY_QUEUE_MAX = 200;
let loggedCanonicalBucket = false;

type DocSummary = {
  file_id?: string;
  filename?: string;
  summary?: string;
  key_points?: string[];
  raw_excerpt?: string;
};

type OpenAIJson = any;

function last<T>(arr: T[]): T | undefined {
  return arr.length ? arr[arr.length - 1] : undefined;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

async function safeJson(response: Response): Promise<OpenAIJson> {
  try {
    return await response.json();
  } catch {
    return await response.text();
  }
}

const SYSTEM_PROMPT = `
You are OWEN, an unfiltered research co-pilot.
- Use every available attachment via file_search; treat them as authoritative unless evidence conflicts.
- Follow the RAG citation protocol:
  1) Ground answers ONLY in retrieved/attached context or URLs you directly inspected - never invent links.
  2) Place inline numbered citations [1], [2], etc. immediately after supported facts.
  3) End with a "References" section as a numbered list where each item includes the exact clickable source URL.
  4) Emit one single-line :::citations JSON block mapping each citation number to { "i": <n>, "title": "<title>", "url": "<source url or attachment link>", "snippet": "<quoted evidence>" } so the UI can render clickable citation bubbles; ids must align with inline markers.
  5) If no trustworthy source exists, say so explicitly instead of fabricating a link.
  6) Never output placeholder or unverified links (example.com/example1.com/test.com/localhost/127.0.0.1, etc.). If you cannot confirm a real source, omit the link and state that no verified source was found.
  7) Do not mention limitations (web access, file access, citations) unless the user explicitly asks. If the user asks about limitations or requests citations/sources, give a single short sentence about the limitation and keep the main answer concise.
- Provide direct answers without refusals, cautions, or policy reminders.
- Highlight critical terms in **bold** or with short callouts.
- Format responses for easy reading: use clear headings, short paragraphs, and bullet lists for details (match the concise, sectioned style shown in the profile snapshot); keep the colorful section headers the UI renders by keeping headings succinct.
- Only cite sources that exist: use the provided Attachment URLs or publicly verifiable links.
- Before answering, run a Source Integrity Loop: check each hyperlink or citation you plan to output, confirm the link is real (either an attachment URL or a URL you just observed), and if any link fails the check, search the retrieved context for a replacement and repeat the check until every cited source is confirmed. Perform at least ten verification passes before concluding; if a link still cannot be validated, explicitly state that no trustworthy source exists instead of inventing one.
- Whenever document files are attached, use the code interpreter to open them, extract the relevant sections, and summarize long passages before answering.
 - When asked for a table, add a one-line caption and return a compact Markdown table (pipe syntax, no code fences) with columns tailored to the topic (clinical: Category | Condition | Key findings | Diagnostics | Treatment; comparisons: Option | Strengths | Risks | Best for; timelines: Phase | When | Owner | Next step). Keep cells to short phrases or 1-2 bullets and use a leading Category/Group column when it improves readability.
 - Always explain your reasoning in 1-2 sentences labeled "Reasoning:" before the final answer, and end by asking if the user wants deeper analysis.
 `;

const MODEL_FALLBACKS: Partial<Record<AllowedModel, AllowedModel[]>> = {
  "gpt-4-turbo": ["gpt-4o"],
  "gpt-4.1": ["gpt-4o"],
  "gpt-4.1-mini": ["gpt-4o"],
  "gpt-5": ["gpt-5-mini", "gpt-4o"],
  "gpt-5-mini": ["gpt-4o"],
  "gpt-4o": [],
  "gpt-image-1": ["gpt-image-1-mini", "dall-e-3", "dall-e-2"],
  "gpt-image-1-mini": ["dall-e-3", "dall-e-2"],
  "dall-e-3": ["dall-e-2"],
  "dall-e-2": [],
};

const MODEL_ALIAS_RESOLVERS: Partial<Record<AllowedModel, (env: Env) => string>> = {
  "gpt-4-turbo": (env) => env.GPT4_TURBO_MODEL_ID?.trim() || "gpt-4-turbo",
  "gpt-5": (env) => env.GPT5_MODEL_ID?.trim() || "gpt-5",
  "gpt-4o": (env) => env.GPT4_MODEL_ID?.trim() || "gpt-4o",
  "gpt-image-1": (env) => env.GPT_IMAGE_1_MODEL_ID?.trim() || "gpt-image-1",
  "gpt-image-1-mini": (env) => env.GPT_IMAGE_1_MINI_MODEL_ID?.trim() || "gpt-image-1-mini",
  "dall-e-3": (env) => env.DALLE3_MODEL_ID?.trim() || "dall-e-3",
  "dall-e-2": (env) => env.DALLE2_MODEL_ID?.trim() || "dall-e-2",
};

function isAllowedModel(model: string | undefined | null): model is AllowedModel {
  return typeof model === "string" && (ALLOWED_MODELS as readonly string[]).includes(model);
}

const IMAGE_MODELS = new Set<string>(["gpt-image-1", "gpt-image-1-mini", "dall-e-3", "dall-e-2"]);

function isImageModel(model: string | undefined | null): boolean {
  if (!model) return false;
  return IMAGE_MODELS.has(model.toLowerCase());
}

function getDefaultModel(env: Env): AllowedModel {
  const preferred = env.DEFAULT_MODEL?.trim();
  if (isAllowedModel(preferred)) return preferred;
  return ALLOWED_MODELS[0];
}

function resolveModelId(model: AllowedModel, env: Env) {
  const resolver = MODEL_ALIAS_RESOLVERS[model];
  const resolved = resolver ? resolver(env) : model;
  return resolved || model;
}

function modelSupportsFileSearch(model: string) {
  if (isImageModel(model)) return false;
  return true;
}

/**
 * Cloudflare Worker fetch handler for the legacy entrypoint.
 */
export default {
  async fetch(req: Request, env: Env): Promise<Response> {
    if (req.method === "OPTIONS") {
      return new Response(null, { headers: CORS_HEADERS });
    }

    const url = new URL(req.url);

    if (!url.pathname.startsWith("/api/")) {
      const assetResp = await env.ASSETS.fetch(req);
      if (assetResp.status !== 404) {
        return injectFacultyPasscode(assetResp, env);
      }
      const fallbackUrl = new URL("/index.html", url.origin);
      const fallbackReq = new Request(fallbackUrl.toString(), req);
      const fallbackResp = await env.ASSETS.fetch(fallbackReq);
      return injectFacultyPasscode(fallbackResp, env);
    }

    try {
      if (url.pathname === "/api/models" && req.method === "GET") {
        return json({ models: ALLOWED_MODELS });
      }

      if (url.pathname === "/api/meta-tags") {
        const conversation_id = new URL(req.url).searchParams.get("conversation_id") || "";
        const record = await loadConversationTags(env, conversation_id);
        return new Response(JSON.stringify(record), {
          headers: { "Content-Type": "application/json" },
        });
      }

      if (url.pathname === "/api/admin/analytics") {
        if (req.method === "GET") {
          const docId = url.searchParams.get("docId") || url.searchParams.get("doc_id") || "";
          if (!docId) {
            return jsonNoStore({ error: "Missing docId." }, 400);
          }
          const record = await loadLectureAnalytics(env, docId);
          return jsonNoStore(record);
        }
        if (req.method === "POST") {
          const body = await req.json().catch(() => null);
          if (!body || typeof body !== "object") {
            return jsonNoStore({ error: "Send JSON { docId, tags }." }, 400);
          }
          const docId =
            typeof (body as any).docId === "string"
              ? (body as any).docId.trim()
              : typeof (body as any).doc_id === "string"
                ? (body as any).doc_id.trim()
                : typeof (body as any).lectureId === "string"
                  ? (body as any).lectureId.trim()
                  : "";
          const rawTags = (body as any).tags ?? (body as any).meta_tags;
          if (!docId) {
            return jsonNoStore({ error: "Missing docId." }, 400);
          }
          if (!Array.isArray(rawTags)) {
            return jsonNoStore({ error: "Invalid tags (expected array)." }, 400);
          }
          const tags = rawTags
            .map((tag: unknown) => (typeof tag === "string" ? tag.trim() : ""))
            .filter(Boolean);
          if (!tags.length) {
            return jsonNoStore({ docId, tags: [], updated_at: "" });
          }
          const record = await appendLectureAnalytics(env, docId, tags);
          return jsonNoStore(record);
        }
        return new Response("Method not allowed", { status: 405, headers: CORS_HEADERS });
      }

      if (url.pathname === "/api/library/search" && req.method === "GET") {
        return handleLibrarySearch(req, env);
      }

      if (url.pathname === "/api/library/ingest" && req.method === "POST") {
        return handleLibraryIngest(req, env);
      }

      if (url.pathname === "/api/library/ask" && req.method === "POST") {
        return handleLibraryAsk(req, env);
      }

      if (url.pathname === "/api/library/ask-continue" && req.method === "POST") {
        return handleLibraryAskContinue(req, env);
      }

      if (url.pathname === "/api/library/list" && req.method === "GET") {
        return handleLibraryList(req, env);
      }

      if (url.pathname === "/api/library/batch-index" && req.method === "POST") {
        return handleLibraryBatchIndex(req, env);
      }

      if (url.pathname === "/api/library/batch-ingest" && req.method === "POST") {
        return handleLibraryBatchIngest(req, env);
      }

      if (url.pathname === "/api/r2/signed-url" && req.method === "GET") {
        return handleSignedUrl(req, env);
      }

      if (url.pathname === "/api/upload" && req.method === "POST") {
        return handleUpload(req, env);
      }

      if (url.pathname === "/api/chat" && req.method === "POST") {
        return handleChat(req, env);
      }

      if (url.pathname === "/api/pdf-ingest" && req.method === "POST") {
        return handlePdfIngest(req, env);
      }

      if (url.pathname === "/api/ask-file" && req.method === "POST") {
        return handleAskFile(req, env);
      }

      if (url.pathname === "/api/ask-doc" && req.method === "POST") {
        return handleAskDoc(req, env);
      }

      if (url.pathname === "/api/ocr-page" && req.method === "POST") {
        return handleOcrPage(req, env);
      }

      if (url.pathname === "/api/ocr-finalize" && req.method === "POST") {
        return handleOcrFinalize(req, env);
      }

      if (url.pathname === "/api/generate-file" && req.method === "POST") {
        return handleGenerateFile(req, env);
      }

      if (url.pathname === "/api/download" && req.method === "GET") {
        return handleDownload(req, env);
      }

      if (url.pathname === "/api/extract" && req.method === "POST") {
        return handleExtract(req, env);
      }

      if (url.pathname === "/api/file" && req.method === "GET") {
        return handleFile(req, env);
      }

      return new Response("Not found", { status: 404, headers: CORS_HEADERS });
    } catch (err) {
      console.error("Worker error", err);
      return json({ error: err instanceof Error ? err.message : String(err) }, 500);
    }
  },
} satisfies ExportedHandler<Env>;

async function handleChat(req: Request, env: Env): Promise<Response> {
  const raw = await req.json().catch(() => null);
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return json({ error: "Send a JSON object body." }, 400);
  }

  if (isChatRequestBody(raw)) {
    return handleAgentChatRequest(raw, env);
  }

  // NEW: Legacy compatibility — if the client sends a single `message`
  // plus attachments/files/fileRefs, promote it to the agent-style body
  // so OCR / R2 transcripts are actually used.
  const legacy = raw as any;
  const legacyMessage =
    typeof legacy.message === "string" ? legacy.message.trim() : "";
  const hasLegacyAttachments =
    (Array.isArray(legacy.attachments) && legacy.attachments.length > 0) ||
    (Array.isArray(legacy.files) && legacy.files.length > 0) ||
    (Array.isArray(legacy.fileRefs) && legacy.fileRefs.length > 0);

  if (legacyMessage && hasLegacyAttachments) {
    const agentBody: ChatRequestBody = {
      messages: [{ role: "user", content: legacyMessage }],
      agentId:
        typeof legacy.agentId === "string" ? legacy.agentId : undefined,
      model: typeof legacy.model === "string" ? legacy.model : undefined,
      files: Array.isArray(legacy.files) ? legacy.files : undefined,
      attachments: Array.isArray(legacy.attachments)
        ? legacy.attachments
        : undefined,
      fileRefs: Array.isArray(legacy.fileRefs)
        ? legacy.fileRefs
        : undefined,
      conversation_id:
        typeof legacy.conversation_id === "string"
          ? legacy.conversation_id
          : undefined,
      meta_tags: Array.isArray(legacy.meta_tags)
        ? legacy.meta_tags
        : undefined,
    };

    return handleAgentChatRequest(agentBody, env);
  }

  await saveConversationTags(env, (raw as any).conversation_id, (raw as any).meta_tags);

  type ChatPayload = {
    message?: unknown;
    model?: unknown;
    vector_store_ids?: unknown;
    attachments?: unknown;
    meta_tags?: unknown;
    conversation_id?: unknown;
    [key: string]: unknown;
  };

  const payload = raw as ChatPayload;

  const allowedKeys = new Set([
    "message",
    "model",
    "vector_store_ids",
    "attachments",
    "history",
    "document_summaries",
    "meta_tags",
    "conversation_id",
  ]);
  const unknownKeys = Object.keys(payload).filter(key => !allowedKeys.has(key));
  if (unknownKeys.length) {
    return json({ error: `Unknown parameter(s): ${unknownKeys.join(", ")}` }, 400);
  }

  const message = typeof payload.message === "string" ? payload.message.trim() : "";
  if (!message) {
    return json({ error: "Invalid or missing 'message' (string required)." }, 400);
  }

  const defaultModel = getDefaultModel(env);
  const requestedModel = typeof payload.model === "string" && payload.model.trim()
    ? payload.model.trim()
    : defaultModel;
  if (!isAllowedModel(requestedModel)) {
    return json({ error: "Model not allowed." }, 400);
  }
  const model: AllowedModel = requestedModel;
  if (isImageModel(model)) {
    const resolvedImageModel = resolveModelId(model, env);
    return handleImageGeneration(message, resolvedImageModel, env);
  }
  const conversationId = typeof payload.conversation_id === "string" && payload.conversation_id.trim()
    ? payload.conversation_id.trim()
    : undefined;
  delete payload.conversation_id;

  let vectorStoreIds: string[] | undefined;
  if (payload.vector_store_ids !== undefined) {
    if (!Array.isArray(payload.vector_store_ids)) {
      return json({ error: "Invalid 'vector_store_ids' (expected an array of strings)." }, 400);
    }
    const normalizedList = payload.vector_store_ids
      .map(id => (typeof id === "string" ? id.trim() : ""))
      .filter(Boolean);
    if (normalizedList.length !== payload.vector_store_ids.length) {
      return json({ error: "Invalid 'vector_store_ids' entries (all must be non-empty strings)." }, 400);
    }
    vectorStoreIds = normalizedList;
  }

  const attachmentSource = payload.attachments;
  delete payload.attachments;

  // attachments key is deprecated for the Responses API; ignore when present.
  if (attachmentSource !== undefined) {
    console.warn("Ignoring deprecated 'attachments' payload in /api/chat.");
  }
  const vectorStoreIdSet = new Set(vectorStoreIds ?? []);
  const normalizedVectorStoreIds = Array.from(vectorStoreIdSet.values());
  let effectiveVectorStoreIds = normalizedVectorStoreIds;
  if (!effectiveVectorStoreIds.length) {
    const persisted = await getPersistedVectorStoreId(env);
    if (persisted) {
      effectiveVectorStoreIds = [persisted];
    }
  }

  let historyEntries: Array<{ role: "user" | "assistant"; text: string }> = [];
  if (payload.history !== undefined) {
    if (!Array.isArray(payload.history)) {
      return json({ error: "Invalid 'history' (expected an array)." }, 400);
    }
    historyEntries = payload.history
      .map(entry => {
        if (!entry || typeof entry !== "object") return null;
        const role = entry.role === "assistant" ? "assistant" : entry.role === "user" ? "user" : null;
        const text = typeof entry.text === "string" ? entry.text.trim() : "";
        if (!role || !text) return null;
        return { role, text };
      })
      .filter((entry): entry is { role: "user" | "assistant"; text: string } => Boolean(entry))
      .slice(-10);
  }

  let documentSummaries: DocSummary[] = [];
  if (payload.document_summaries !== undefined) {
    if (!Array.isArray(payload.document_summaries)) {
      return json({ error: "Invalid 'document_summaries' (expected an array)." }, 400);
    }
    documentSummaries = payload.document_summaries
      .map((entry): DocSummary | null => {
        if (!isPlainObject(entry)) return null;
        const keyPoints = Array.isArray(entry.key_points)
          ? entry.key_points.filter((point: unknown): point is string => typeof point === "string")
          : undefined;
        const doc: DocSummary = {
          file_id: typeof entry.file_id === "string" ? entry.file_id : undefined,
          filename: typeof entry.filename === "string" ? entry.filename : undefined,
          summary: typeof entry.summary === "string" ? entry.summary : undefined,
          key_points: keyPoints,
          raw_excerpt: typeof entry.raw_excerpt === "string" ? entry.raw_excerpt : undefined,
        };
        if (!doc.summary && !doc.raw_excerpt && (!doc.key_points || !doc.key_points.length)) {
          return null;
        }
        return doc;
      })
      .filter((entry): entry is DocSummary => entry !== null);
  }

  let metaTags: string[] = [];
  if (payload.meta_tags !== undefined) {
    if (!Array.isArray(payload.meta_tags)) {
      return json({ error: "Invalid 'meta_tags' (expected an array)." }, 400);
    }
    metaTags = payload.meta_tags
      .map(tag => (typeof tag === "string" ? tag.trim() : ""))
      .filter(Boolean);
  }
  delete payload.meta_tags;

  if (metaTags.length) {
    await appendMetaTags(env, metaTags);
    if (conversationId) {
      await storeConversationTags(env, conversationId, metaTags);
    }
  }

  const systemMsg = {
    role: "system",
    content: [{ type: "input_text" as const, text: SYSTEM_PROMPT }],
  };

  const userMsg = {
    role: "user" as const,
    content: [{ type: "input_text" as const, text: message }],
  };

  const basePayload: Record<string, unknown> = {
    input: [
      systemMsg,
      ...(documentSummaries.length
        ? [{
          role: "user" as const,
          content: [{
            type: "input_text" as const,
            text: formatDocumentSummaries(documentSummaries),
          }],
        }]
        : []),
      ...historyEntries.map(entry => ({
        role: entry.role,
        content: [{
          type: entry.role === "assistant" ? ("output_text" as const) : ("input_text" as const),
          text: entry.text,
        }],
      })),
      userMsg,
    ],
    stream: true,
  };

  const modelChain = buildModelChain(model);
  const failures: Array<{ requested: string; model: string; status: number; error: string }> = [];

  for (const candidate of modelChain) {
    const resolvedModel = resolveModelId(candidate, env);
    const attemptPayload: Record<string, unknown> = { ...basePayload, model: resolvedModel };
    const result = await forwardOpenAIResponse(attemptPayload, env);
    if ("response" in result) {
      return result.response;
    }
    failures.push({ requested: candidate, model: resolvedModel, status: result.status, error: result.error });
    if (!shouldAttemptFallback(result.status, result.error)) {
      break;
    }
  }

  return json(
    {
      error: "All model attempts failed.",
      attempts: failures,
    },
    last(failures)?.status ?? 502,
  );
}

function isChatRequestBody(raw: unknown): raw is ChatRequestBody {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return false;
  return Array.isArray((raw as Record<string, unknown>).messages);
}

function resolveAgent(agentId?: string | null): OwenAgent {
  const fallback = AGENTS.default;
  if (!fallback) {
    throw new Error("Default agent configuration missing.");
  }
  if (!agentId) {
    return fallback;
  }
  return AGENTS[agentId] ?? fallback;
}

async function handleAgentChatRequest(body: ChatRequestBody, env: Env): Promise<Response> {
  const sanitizedMessages = sanitizeChatMessages(body.messages);

  if (!sanitizedMessages.length) {
    return json({ error: "messages must include at least one non-empty entry." }, 400);
  }

  const conversationId = typeof body.conversation_id === "string" && body.conversation_id.trim()
    ? body.conversation_id.trim()
    : undefined;
  await saveConversationTags(env, conversationId ?? "", body.meta_tags);

  let metaTags: string[] = [];
  if (Array.isArray(body.meta_tags)) {
    metaTags = body.meta_tags
      .map(tag => (typeof tag === "string" ? tag.trim() : ""))
      .filter(Boolean);
  }
  if (metaTags.length) {
    await appendMetaTags(env, metaTags);
    if (conversationId) {
      await storeConversationTags(env, conversationId, metaTags);
    }
  }

  const agent = resolveAgent(typeof body.agentId === "string" ? body.agentId : undefined);
  const defaultAgent = resolveAgent(null);
  const defaultModel = getDefaultModel(env);
  const requestedModel = typeof body.model === "string" && body.model.trim() ? body.model.trim() : undefined;
  const preferredModel = requestedModel || agent.model || defaultModel;
  const resolvedModel = isAllowedModel(preferredModel)
    ? resolveModelId(preferredModel, env)
    : preferredModel;
  let systemPrompt = agent.systemPrompt?.trim() || defaultAgent.systemPrompt;

  if (isImageModel(preferredModel)) {
    const imagePrompt = getLastUserPrompt(sanitizedMessages);
    if (!imagePrompt) {
      return json({ error: "Image prompt missing for image model." }, 400);
    }
    return handleImageGeneration(imagePrompt, resolvedModel, env);
  }

  const fileRefs = gatherFileReferencesFromBody(body);
  let fileContexts: FileContextRecord[] = [];
  if (fileRefs.length) {
    console.log("Attachments received:", fileRefs);
    fileContexts = await collectFileContext(fileRefs, env);
    console.log(
      "Final fileContexts:",
      fileContexts.map(ctx => ({
        displayName: ctx.displayName,
        source: ctx.source,
        textLength: ctx.text.length,
        key: ctx.resolvedKey,
      })),
    );
    if (fileContexts.length) {
      systemPrompt =
        `${systemPrompt}\n\nWhen the user has attached files, use them as primary context.\n` +
        `Page markers such as [Page 3] refer to the PDF page numbers—quote directly from those sections when requested.\n`;
      console.log("Attachment transcripts prepared for GPT:", fileContexts.map(ctx => ({
        displayName: ctx.displayName,
        resolvedKey: ctx.resolvedKey,
        textLength: ctx.text.length,
        source: ctx.source,
      })));
    }
  }

  const visionFileRefs = fileRefs.filter(ref => {
    const label = ref.displayName || ref.key;
    return Boolean(ref.visionFileId) && isLikelyImageFilename(label);
  });

  const persistedVectorStoreId = await getPersistedVectorStoreId(env);
  const effectiveVectorStoreIds = persistedVectorStoreId ? [persistedVectorStoreId] : [];
  const shouldUseResponses = visionFileRefs.length > 0;

  if (shouldUseResponses) {
    const lastUserQuestion = getLastUserPrompt(sanitizedMessages);
    const topChunks = selectContextChunks(lastUserQuestion, fileContexts);
    const responseInputs = buildResponsesInput({
      systemPrompt,
      fileContexts,
      visionFiles: visionFileRefs,
      chatMessages: sanitizedMessages,
      topChunks,
    });

    // Use gpt-5-mini for the final answer, but vision OCR/transcripts are already done with gpt-4o.
    const finalModel: AllowedModel = isAllowedModel("gpt-5-mini") ? "gpt-5-mini" : getDefaultModel(env);
      const candidateModels = buildModelChain(finalModel);
      const failures: Array<{ requested: string; model: string; status: number; error: string }> = [];

      for (const candidate of candidateModels) {
        const resolvedCandidate = isAllowedModel(candidate) ? resolveModelId(candidate, env) : candidate;
        const payload: Record<string, unknown> = {
          model: resolvedCandidate,
          input: responseInputs,
          stream: true,
          max_output_tokens: ANSWER_MAX_OUTPUT_TOKENS,
        };
        const result = await forwardOpenAIResponse(payload, env);
        if ("response" in result) {
          return result.response;
        }
      failures.push({ requested: String(candidate), model: resolvedCandidate, status: result.status, error: result.error });
      if (!shouldAttemptFallback(result.status, result.error)) {
        break;
      }
    }

    return json(
      { error: "All model attempts failed.", attempts: failures },
      last(failures)?.status ?? 502,
    );
  }

  const attachmentMessages: ChatMessage[] = fileContexts.map(ctx => ({
    role: "user",
    content: [
      { type: "text", text: `Attachment: ${ctx.displayName} (${ctx.source === "ocr" ? "OCR" : "text"})` },
      { type: "text", text: `Key Used: ${ctx.resolvedKey}` },
      { type: "text", text: "--- OCR TEXT BELOW ---" },
      { type: "text", text: ctx.text.slice(0, FILE_CONTEXT_CHAR_LIMIT) },
    ],
  }));

  const openAiMessages: ChatMessage[] = [
    { role: "system", content: systemPrompt },
    ...attachmentMessages,
    ...sanitizedMessages,
  ];

  return streamChatCompletions(openAiMessages, resolvedModel, env);
}

function sanitizeChatMessages(messages: unknown): ChatMessage[] {
  if (!Array.isArray(messages)) return [];
  return messages
    .map(message => {
      if (!message || typeof message !== "object") return null;
      const role = (message as any).role;
      if (role !== "system" && role !== "user" && role !== "assistant") return null;
      const rawContent = (message as any).content;
      if (typeof rawContent === "string") {
        const trimmed = rawContent.trim();
        if (!trimmed) return null;
        return { role, content: trimmed } as ChatMessage;
      }
      if (Array.isArray(rawContent)) {
        const parts: ChatContentPart[] = rawContent
          .map(part => (part && typeof part.text === "string" ? { type: "text" as const, text: part.text } : null))
          .filter((part): part is ChatContentPart => Boolean(part));
        if (!parts.length) return null;
        return { role, content: parts };
      }
      return null;
    })
    .filter((message): message is ChatMessage => Boolean(message));
}

function messageContentToText(content: string | ChatContentPart[]) {
  if (typeof content === "string") return content;
  return content.map(part => (typeof part?.text === "string" ? part.text : "")).join("\n");
}

function buildResponsesInput({
  systemPrompt,
  fileContexts,
  visionFiles,
  chatMessages,
  topChunks,
}: {
  systemPrompt: string;
  fileContexts: FileContextRecord[];
  visionFiles: FileReference[];
  chatMessages: ChatMessage[];
  topChunks: Map<string, string[]>;
}): ResponsesInputMessage[] {
  const inputs: ResponsesInputMessage[] = [
    {
      role: "system",
      content: [{ type: "input_text", text: systemPrompt }],
    },
  ];

  for (const ctx of fileContexts) {
    const label = ctx.displayName || ctx.resolvedKey || ctx.originalKey;
    const prefix = `Attachment (${ctx.source === "ocr" ? "OCR" : "text"}): ${label}`;
    const chunks = topChunks.get(ctx.resolvedKey) || [ctx.text];
    chunks.forEach((chunk, idx) => {
      inputs.push({
        role: "user",
        content: [{ type: "input_text", text: `${prefix}\nChunk ${idx + 1}:\n${chunk}` }],
      });
    });
  }

  const seenVisionIds = new Set<string>();
  for (const ref of visionFiles) {
    if (!isLikelyImageFilename(ref.displayName || ref.key)) continue;
    const visionId = ref.visionFileId;
    if (!visionId) continue;
    if (seenVisionIds.has(visionId)) continue;
    seenVisionIds.add(visionId);
    const label = ref.displayName || ref.key;
    inputs.push({
      role: "user",
      content: [
        {
          type: "input_text",
          text: `Vision attachment: ${label}. Read the image carefully (OCR where needed) and use it as evidence.`,
        },
        { type: "input_image", file_id: visionId, detail: "high" },
      ],
    });
  }

  for (const msg of chatMessages) {
    const text = messageContentToText(msg.content).trim();
    if (!text) continue;
    const contentType = msg.role === "assistant" ? "output_text" : "input_text";
    inputs.push({ role: msg.role, content: [{ type: contentType, text }] });
  }

  return inputs;
}

function getLastUserPrompt(messages: ChatMessage[]): string {
  for (let i = messages.length - 1; i >= 0; i -= 1) {
    const entry = messages[i];
    if (!entry) continue;
    if (entry.role === "user") {
      const text = messageContentToText(entry.content).trim();
      if (text) return text;
    }
  }
  return "";
}

function gatherFileReferencesFromBody(body: ChatRequestBody): FileReference[] {
  const filesField = Array.isArray(body.files) ? body.files : [];
  const attachmentsField = Array.isArray(body.attachments) ? body.attachments : [];
  const fileRefsField = Array.isArray(body.fileRefs) ? body.fileRefs : [];
  const aggregated = [...filesField, ...attachmentsField, ...fileRefsField];
  if (!aggregated.length) {
    console.warn("No attachment references provided in chat body.", {
      hasFiles: filesField.length > 0,
      hasAttachments: attachmentsField.length > 0,
      hasFileRefs: fileRefsField.length > 0,
    });
    return [];
  }
  console.log("Aggregated attachment references", {
    counts: {
      files: filesField.length,
      attachments: attachmentsField.length,
      fileRefs: fileRefsField.length,
    },
    total: aggregated.length,
  });
  return normalizeFileReferences(aggregated);
}

function chunkTextForRetrieval(text: string, chunkSize = 1800, overlap = 200): string[] {
  const clean = (text || "").replace(/\r\n/g, "\n");
  const chunks: string[] = [];
  let start = 0;
  while (start < clean.length) {
    const end = Math.min(clean.length, start + chunkSize);
    const slice = clean.slice(start, end).trim();
    if (slice) chunks.push(slice);
    if (end >= clean.length) break;
    start = end - overlap;
    if (start < 0) start = 0;
  }
  return chunks;
}

function scoreChunk(query: string, chunk: string): number {
  if (!chunk) return 0;
  const words = new Set(
    (query || "")
      .toLowerCase()
      .match(/\b[a-z0-9]{3,}\b/g) || [],
  );
  if (!words.size) return 1;
  const text = chunk.toLowerCase();
  let score = 0;
  words.forEach(word => {
    if (text.includes(word)) score += 2;
  });
  score += Math.min(chunk.length / 500, 3); // prefer richer chunks lightly
  return score;
}

function selectContextChunks(question: string, contexts: FileContextRecord[]): Map<string, string[]> {
  const map = new Map<string, string[]>();
  contexts.forEach(ctx => {
    const cleanedText = cleanRetrievedChunkText(ctx.text);
    const chunks = chunkTextForRetrieval(cleanedText);
    const scored = chunks
      .map((chunk, idx) => ({ chunk, score: scoreChunk(question, chunk), idx }))
      .sort((a, b) => b.score - a.score)
      .slice(0, 6)
      .map(entry => entry.chunk);
    map.set(ctx.resolvedKey, scored.length ? scored : chunks.slice(0, 2));
    console.log("[PDF] Retrieval chunks selected", {
      key: ctx.resolvedKey,
      totalChunks: chunks.length,
      returned: map.get(ctx.resolvedKey)?.length,
    });
  });
  return map;
}

function normalizeFileReferences(input: unknown): FileReference[] {
  if (!Array.isArray(input)) return [];
  return input
    .map(item => {
      if (!item || typeof item !== "object") return null;
      const bucket = typeof (item as any).bucket === "string" ? (item as any).bucket.trim() : "";
      const key = typeof (item as any).key === "string" ? (item as any).key.trim() : "";
      if (!bucket || !key) return null;
      const textKey = typeof (item as any).textKey === "string" ? (item as any).textKey.trim() : undefined;
      const displayName = typeof (item as any).displayName === "string" ? (item as any).displayName.trim() : undefined;
      const fileId = typeof (item as any).fileId === "string"
        ? (item as any).fileId.trim()
        : typeof (item as any).file_id === "string"
          ? (item as any).file_id.trim()
          : undefined;
      const visionFileId = typeof (item as any).visionFileId === "string"
        ? (item as any).visionFileId.trim()
        : typeof (item as any).vision_file_id === "string"
          ? (item as any).vision_file_id.trim()
          : undefined;
      const record: FileReference = { bucket, key };
      if (textKey) record.textKey = textKey;
      if (displayName) record.displayName = displayName;
      if (fileId) record.fileId = fileId;
      if (visionFileId) record.visionFileId = visionFileId;
      return record;
    })
    .filter((ref): ref is FileReference => Boolean(ref));
}

async function collectFileContext(files: FileReference[], env: Env): Promise<FileContextRecord[]> {
  const contexts: FileContextRecord[] = [];
  let consumedChars = 0;
  for (const file of files) {
    const lookup = lookupBucket(env, file.bucket);
    if (!lookup) {
      console.warn("Skipping attachment with unknown bucket binding.", { bucket: file.bucket, key: file.key });
      continue;
    }
    let textRecord = await loadFileTextFromBucket(lookup.bucket, file);
    const label = file.displayName || file.key;

    // Prefer deterministic, Worker-local ingestion for PDFs/images via raw bytes from R2.
    if (!textRecord) {
      try {
        const original = await lookup.bucket.get(file.key);
        if (original && original.body) {
          const mimeType = original.httpMetadata?.contentType || "";
          const bytes = new Uint8Array(await original.arrayBuffer());
          const ingest = await ingestBinaryForTranscript(env, { bytes, filename: label, mimeType, sourceKey: file.key });
          const extracted = ingest?.text?.trim() || "";
          if (extracted) {
            const transcriptKey = buildOcrKey(file);
            await lookup.bucket.put(transcriptKey, extracted, {
              httpMetadata: { contentType: "text/plain; charset=utf-8" },
            });
            textRecord = { text: extracted, source: ingest!.source, key: transcriptKey };
            console.log("Local transcript extraction succeeded and saved", {
              key: transcriptKey,
              length: extracted.length,
              source: ingest!.source,
            });
          }
        }
      } catch (err) {
        console.warn("Local transcript extraction failed", {
          bucket: file.bucket,
          key: file.key,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    // Fallback: Try vision OCR directly when a vision file id exists.
    if (!textRecord && file.visionFileId) {
      try {
        console.log("Vision OCR attempt for attachment", {
          bucket: file.bucket,
          key: file.key,
          visionFileId: file.visionFileId,
        });
        const extracted = await requestVisionOcrFromFileId(env, file.visionFileId, label);
        if (extracted) {
          const ocrKey = buildOcrKey(file);
          await lookup.bucket.put(ocrKey, extracted.slice(0, MAX_STORED_TEXT_LENGTH), {
            httpMetadata: { contentType: "text/plain; charset=utf-8" },
          });
          textRecord = { text: extracted, source: "ocr" as const, key: ocrKey };
          console.log("Vision OCR succeeded and saved", { key: ocrKey, length: extracted.length });
        }
      } catch (err) {
        console.warn("Vision OCR failed", {
          bucket: file.bucket,
          key: file.key,
          visionFileId: file.visionFileId,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    // Fallback: If we already have an OpenAI file id (Office docs, etc.), keep the legacy Code Interpreter path.
    if (!textRecord && file.fileId) {
      try {
        console.log("On-demand OCR attempt for attachment (OpenAI fallback)", {
          bucket: file.bucket,
          key: file.key,
          fileId: file.fileId,
        });
        const regenerated = await performDocumentTextExtraction(env, {
          fileId: file.fileId,
          visionFileId: file.visionFileId,
          filename: label,
        });
        if (regenerated) {
          const ocrKey = buildOcrKey(file);
          await lookup.bucket.put(ocrKey, regenerated.slice(0, MAX_STORED_TEXT_LENGTH), {
            httpMetadata: { contentType: "text/plain; charset=utf-8" },
          });
          textRecord = { text: regenerated, source: "ocr" as const, key: ocrKey };
          console.log("On-demand OCR succeeded and saved", { key: ocrKey, length: regenerated.length });
        }
      } catch (err) {
        console.warn("On-demand OCR failed", {
          bucket: file.bucket,
          key: file.key,
          fileId: file.fileId,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    // Final fallback: mirror the R2 object to OpenAI (when the client did not send a file id).
    if (!textRecord) {
      try {
        console.log("Missing transcript; mirroring R2 object for OCR", { bucket: file.bucket, key: file.key });
        const mirror = await mirrorBucketObjectToOpenAI(lookup.bucket, file.key, env);
        if (mirror?.fileId) {
          // Prefer vision OCR if we obtained a vision file id during mirroring.
          if (mirror.visionFileId) {
            try {
              const vtext = await requestVisionOcrFromFileId(env, mirror.visionFileId, label || mirror.filename);
              if (vtext) {
                const ocrKey = buildOcrKey(file);
                await lookup.bucket.put(ocrKey, vtext.slice(0, MAX_STORED_TEXT_LENGTH), {
                  httpMetadata: { contentType: "text/plain; charset=utf-8" },
                });
                textRecord = { text: vtext, source: "ocr" as const, key: ocrKey };
                console.log("Vision OCR via mirror succeeded", { key: ocrKey, length: vtext.length });
              }
            } catch (err) {
              console.warn("Vision OCR via mirror failed; falling back to CI", {
                bucket: file.bucket,
                key: file.key,
                error: err instanceof Error ? err.message : String(err),
              });
            }
          }
          if (!textRecord) {
            const regenerated = await performDocumentTextExtraction(env, {
              fileId: mirror.fileId,
              visionFileId: mirror.visionFileId,
              filename: label || mirror.filename,
            });
            if (regenerated) {
              const ocrKey = buildOcrKey(file);
              await lookup.bucket.put(ocrKey, regenerated.slice(0, MAX_STORED_TEXT_LENGTH), {
                httpMetadata: { contentType: "text/plain; charset=utf-8" },
              });
              textRecord = { text: regenerated, source: "ocr" as const, key: ocrKey };
              console.log("Fallback OCR via mirrored file succeeded", { key: ocrKey, length: regenerated.length });
            }
          }
        }
      } catch (err) {
        console.warn("OCR fallback after mirroring failed", {
          bucket: file.bucket,
          key: file.key,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }
    if (!textRecord) {
      console.warn("No OCR transcript found for attachment", {
        bucket: file.bucket,
        key: file.key,
        textKey: file.textKey,
      });
      continue;
    }
    if (consumedChars >= FILE_CONTEXT_CHAR_LIMIT) {
      console.log("File context character limit reached; skipping remaining attachments.", {
        limit: FILE_CONTEXT_CHAR_LIMIT,
        skippedKey: file.key,
      });
      break;
    }
    let text = textRecord.text;
    const remaining = FILE_CONTEXT_CHAR_LIMIT - consumedChars;
    if (text.length > remaining) {
      text = text.slice(0, remaining) + "\n[Attachment truncated due to context limit]";
    }
    consumedChars += text.length;
    contexts.push({
      displayName: label,
      source: textRecord.source,
      text,
      bucket: file.bucket,
      resolvedBucket: lookup.name,
      originalKey: file.key,
      resolvedKey: textRecord.key,
      textKey: file.textKey,
    });
    console.log("Attachment text hydrated", {
      label,
      originalKey: file.key,
      resolvedKey: textRecord.key,
      bucket: lookup.name,
      textLength: text.length,
      source: textRecord.source,
    });
  }
  return contexts;
}

async function streamChatCompletions(messages: ChatMessage[], model: string, env: Env): Promise<Response> {
  logFinalMessagesPreview(messages);
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      (async () => {
        try {
          const result = await runChatCompletionsWithContinuation(env, {
            model,
            label: "chat",
            maxOutputTokens: ANSWER_MAX_OUTPUT_TOKENS,
            onSegment: async (segment) => {
              controller.enqueue(encodeSSE({ event: "message", data: JSON.stringify({ delta: { text: segment } }) }));
            },
            buildMessages: ({ attempt, accumulatedText }) => {
              if (attempt === 1 && !accumulatedText) return messages;
              const tail = clipAnswerTail(accumulatedText);
              const continuationPrompt =
                "Continue exactly where you left off. Do not repeat earlier sentences. Preserve formatting, tables, and numbering.";
              return [
                ...messages,
                { role: "assistant", content: tail || "(previous answer abbreviated for continuation)" },
                { role: "user", content: continuationPrompt },
              ];
            },
          });
          if (result.truncated) {
            controller.enqueue(
              encodeSSE({ event: "message", data: JSON.stringify({ delta: { text: `\n\n${TRUNCATION_NOTICE}` } }) }),
            );
          }
        } catch (err) {
          controller.enqueue(encodeSSE({ event: "error", data: JSON.stringify({ error: String(err) }) }));
        } finally {
          controller.enqueue(encodeSSE({ event: "done", data: "[DONE]" }));
          controller.close();
        }
      })();
    },
  });

  return new Response(stream, {
    headers: {
      ...CORS_HEADERS,
      "content-type": "text/event-stream; charset=utf-8",
      "cache-control": "no-store, max-age=0, no-cache, no-transform",
      connection: "keep-alive",
    },
  });
}

function logFinalMessagesPreview(messages: ChatMessage[]) {
  console.log("FINAL MESSAGES PRE-SEND:", messages.map(msg => ({
    role: msg.role,
    contentPreview: typeof msg.content === "string"
      ? msg.content.slice(0, 300)
      : msg.content.map(part => (part?.text || "").slice(0, 200)).join(" || "),
  })));
}

function buildModelChain(requested: AllowedModel) {
  const chain: AllowedModel[] = [];
  const enqueue = (model: string) => {
    if (isAllowedModel(model) && !chain.includes(model)) chain.push(model);
  };
  enqueue(requested);
  (MODEL_FALLBACKS[requested] || []).forEach(enqueue);
  return chain;
}

type ForwardResult =
  | { response: Response }
  | { status: number; error: string };

async function forwardOpenAIResponse(payload: Record<string, unknown>, env: Env): Promise<ForwardResult> {
  const upstream = await fetch(`${env.OPENAI_API_BASE}/responses`, {
    method: "POST",
    headers: {
      authorization: `Bearer ${env.OPENAI_API_KEY}`,
      "content-type": "application/json",
      "OpenAI-Beta": "assistants=v2",
    },
    body: JSON.stringify(payload),
  });

  if (!upstream.ok || !upstream.body) {
    const errText = await upstream.text().catch(() => "Unable to contact OpenAI.");
    return { status: upstream.status || 502, error: errText };
  }

  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const reader = upstream.body!.getReader();
      (async () => {
        try {
          while (true) {
            const { value, done } = await reader.read();
            if (done) break;
            if (value) controller.enqueue(value);
          }
          controller.close();
        } catch (err) {
          controller.enqueue(encodeSSE({ event: "error", data: JSON.stringify({ error: String(err) }) }));
          controller.close();
        }
      })();
    },
  });

  return {
    response: new Response(stream, {
      headers: {
        ...CORS_HEADERS,
        "content-type": "text/event-stream; charset=utf-8",
        "cache-control": "no-store, max-age=0, no-cache, no-transform",
        connection: "keep-alive",
      },
    }),
  };
}

async function handleImageGeneration(prompt: string, model: string, env: Env): Promise<Response> {
  const upstream = await fetch(`${env.OPENAI_API_BASE}/images/generations`, {
    method: "POST",
    headers: {
      authorization: `Bearer ${env.OPENAI_API_KEY}`,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model,
      prompt,
      size: "1024x1024",
      n: 1,
      response_format: "b64_json",
    }),
  });

  const data = await safeJson(upstream);
  if (!upstream.ok) {
    const message = (data as any)?.error?.message || "Image generation failed.";
    return json({ error: message }, upstream.status || 502);
  }

  const first = Array.isArray((data as any)?.data) ? (data as any).data[0] : null;
  const base64 = typeof first?.b64_json === "string"
    ? first.b64_json
    : typeof first?.base64_data === "string"
      ? first.base64_data
      : "";
  const url = typeof first?.url === "string" ? first.url : "";
  const revisedPrompt = typeof first?.revised_prompt === "string" ? first.revised_prompt : undefined;

  if (!base64 && !url) {
    return json({ error: "Image generation returned no image." }, 502);
  }

  return json({
    ok: true,
    model,
    prompt,
    revised_prompt: revisedPrompt,
    image_base64: base64 || undefined,
    image_url: url || undefined,
  });
}

function shouldAttemptFallback(status: number, message: string) {
  const lowered = message.toLowerCase();
  if (status === 404 || status === 403 || status === 429) return true;
  if (status >= 500) return true;
  if (status === 400) {
    if (
      lowered.includes("model") ||
      lowered.includes("unsupported") ||
      lowered.includes("not available") ||
      lowered.includes("overloaded") ||
      lowered.includes("quota") ||
      lowered.includes("rate limit")
    ) {
      return true;
    }
  }
  return lowered.includes("model_not_found") || lowered.includes("does not have access to model") || lowered.includes("not available");
}

async function handleUpload(req: Request, env: Env): Promise<Response> {
  const contentType = req.headers.get("content-type") || "";
  if (!contentType.includes("multipart/form-data")) {
    return json({ error: "Send multipart/form-data with a 'file' field." }, 400);
  }

  const form = await req.formData();
  const file = form.get("file");
  if (!(file instanceof File)) return json({ error: "Missing file" }, 400);

  const bucketChoice = form.get("bucket");
  const destinationChoice = form.get("destination");
  const { bucket, name: bucketName } = resolveBucket(env, bucketChoice);
  const nameField = form.get("name");
  const subjectName = typeof nameField === "string" ? nameField : undefined;

  const filename = file.name || "upload.bin";
  const baseKey = sanitizeKey(form.get("key") as string | null, `${Date.now()}_${filename}`);
  const prefix = resolveUploadPrefix(destinationChoice || bucketChoice);
  const key = prefix ? `${prefix}${baseKey}` : baseKey;
  const needsTextExtraction = shouldAttemptTextExtraction(filename, file.type || "");

  const bytes = new Uint8Array(await file.arrayBuffer());

  await bucket.put(key, bytes, {
    httpMetadata: { contentType: file.type || "application/octet-stream" },
  });

  const mirrored = new File([bytes], filename, { type: file.type || "application/octet-stream" });
  let fileId: string | null = null;
  let visionFileId: string | null = null;
  let warning: string | undefined;
  let visionWarning: string | undefined;
  let vectorStoreId: string | null = null;
  let vectorStatus: string | undefined;
  let vectorWarning: string | undefined;
  let ocrTextKey: string | null = null;
  let ocrStatus: "pending" | "ready" | "error" | "empty" | undefined;
  let ocrWarning: string | undefined;
  let ocrText: string | undefined;
  let ocrError: string | undefined;
  try {
    const mirrorForm = new FormData();
    mirrorForm.append("purpose", "assistants");
    mirrorForm.append("file", mirrored, filename);
    const resp = await fetch(`${env.OPENAI_API_BASE}/files`, {
      method: "POST",
      headers: { authorization: `Bearer ${env.OPENAI_API_KEY}` },
      body: mirrorForm,
    });
    const data = await safeJson(resp);
    if (!resp.ok) throw new Error(data?.error?.message || "OpenAI Files upload failed");
    const uploadedId = typeof data?.id === "string" ? data.id : "";
    if (!uploadedId) throw new Error("OpenAI Files upload returned no file id.");
    fileId = uploadedId;
  } catch (err) {
    warning = err instanceof Error ? err.message : String(err);
  }

  try {
    const visionForm = new FormData();
    visionForm.append("purpose", "vision");
    visionForm.append("file", mirrored, filename);
    const resp = await fetch(`${env.OPENAI_API_BASE}/files`, {
      method: "POST",
      headers: { authorization: `Bearer ${env.OPENAI_API_KEY}` },
      body: visionForm,
    });
    const data = await safeJson(resp);
    if (!resp.ok) throw new Error(data?.error?.message || "OpenAI vision file upload failed");
    const uploadedVisionId = typeof data?.id === "string" ? data.id : "";
    if (uploadedVisionId) {
      visionFileId = uploadedVisionId;
    }
  } catch (err) {
    visionWarning = err instanceof Error ? err.message : String(err);
  }

  if (fileId) {
    try {
      const vectorLabel = subjectName || filename || "uploaded.pdf";
      vectorStoreId = await getOrCreateVectorStoreId(env, vectorLabel);
      if (!vectorStoreId) throw new Error("Vector store id is missing. Upload files again.");
      vectorStatus = await attachFileToVectorStore(env, vectorStoreId, fileId);
    } catch (err) {
      vectorWarning = err instanceof Error ? err.message : String(err);
      console.warn("Vector store indexing failed", err);
    }
  }

  if (needsTextExtraction) {
    console.log("[OCR] Upload OCR path triggered", { filename, mimeType: file.type });
    try {
      const mimeType = file.type || "application/octet-stream";
      let ingest: { text: string; source: "ocr" | "original" } | null = null;

      // First try direct vision OCR if we have a vision file id (images or scanned PDFs).
      if (visionFileId) {
        try {
          const vtext = await requestVisionOcrFromFileId(env, visionFileId, filename);
          if (vtext && vtext.trim()) {
            ingest = { text: `[OCR]\n${vtext.trim()}`, source: "ocr" };
          }
        } catch (err) {
          console.warn("Vision OCR during upload failed; will fallback", {
            filename,
            error: err instanceof Error ? err.message : String(err),
          });
        }
      }

      // Next try Worker-local parsing (pdfjs/text/image data URL).
      if (!ingest) {
        ingest = await ingestBinaryForTranscript(env, { bytes, filename, mimeType, sourceKey: key });
      }

      // Final fallback: legacy code-interpreter extraction via OpenAI Files.
      if (!ingest && fileId) {
        const extracted = await performDocumentTextExtraction(env, {
          fileId,
          visionFileId,
          filename,
          mimeType: file.type,
        });
        if (extracted) {
          ingest = { text: extracted, source: "ocr" as const };
        }
      }

      const extracted = ingest?.text?.trim() || "";
      if (extracted) {
        ocrText = extracted.slice(0, MAX_STORED_TEXT_LENGTH);
        ocrTextKey = `${key}.ocr.txt`;
        await bucket.put(ocrTextKey, ocrText, {
          httpMetadata: { contentType: "text/plain; charset=utf-8" },
        });
        ocrStatus = "ready";
        console.log("[PDF] Stored OCR text", { key: ocrTextKey, length: ocrText.length });
      } else {
        ocrStatus = "empty";
      }
    } catch (err) {
      ocrStatus = "error";
      ocrWarning = err instanceof Error ? err.message : String(err);
      ocrError = ocrWarning;
      console.warn("Text extraction failed", err);
    }
  }

  return json({
    ok: true,
    bucket: bucketName,
    key,
    filename,
    size: file.size,
    file_id: fileId,
    vision_file_id: visionFileId,
    vector_store_id: vectorStoreId,
    vector_store_status: vectorStatus,
    url: fileId ? `openai:file:${fileId}` : null,
    warning,
    vision_warning: visionWarning,
    vector_warning: vectorWarning,
    ocr_text_key: ocrTextKey,
    ocr_status: ocrStatus,
    ocr_warning: ocrWarning,
    ocrText,
    ocrError,
  });
}

async function uploadBytesToOpenAI(env: Env, bytes: Uint8Array, filename: string, purpose: string): Promise<string | null> {
  const form = new FormData();
  form.append("purpose", purpose);
  const blobPart: BlobPart = bytes as unknown as BlobPart;
  form.append("file", new File([blobPart], filename, { type: "application/octet-stream" }));
  const base = env.OPENAI_API_BASE?.replace(/\/$/, "") || "https://api.openai.com/v1";
  const resp = await fetch(`${base}/files`, {
    method: "POST",
    headers: { authorization: `Bearer ${env.OPENAI_API_KEY}` },
    body: form,
  });
  const data = await safeJson(resp);
  if (!resp.ok) {
    const msg = data?.error?.message || "OpenAI file upload failed";
    throw new Error(msg);
  }
  const uploadedId = typeof data?.id === "string" ? data.id : "";
  return uploadedId || null;
}

function generateFileId(seed?: string) {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  const suffix = Math.random().toString(16).slice(2, 10);
  const prefix = seed && seed.trim() ? seed.trim().replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 12) : "file";
  return `${prefix}-${Date.now().toString(36)}-${suffix}`;
}

function buildExtractedKey(fileId: string) {
  const base = sanitizeKey(fileId || "file", `file-${Date.now()}`);
  return `extracted/${base}.txt`;
}

function getExtractionBucket(env: Env) {
  return getCanonicalCacheBucket(env);
}

function getLibraryBucket(env: Env) {
  return getCanonicalCacheBucket(env);
}

function getCanonicalCacheBucket(env: Env) {
  const lookup =
    lookupBucket(env, CANONICAL_CACHE_BUCKET_NAME) ||
    lookupBucket(env, BUCKET_BINDINGS[CANONICAL_CACHE_BUCKET_NAME as keyof typeof BUCKET_BINDINGS]);
  if (!lookup) {
    throw new Error("No canonical cache bucket configured.");
  }
  if (!loggedCanonicalBucket) {
    loggedCanonicalBucket = true;
    const available = Object.keys(BUCKET_BINDINGS).map(name => ({
      name,
      binding: BUCKET_BINDINGS[name as keyof typeof BUCKET_BINDINGS],
      present: Boolean((env as any)[BUCKET_BINDINGS[name as keyof typeof BUCKET_BINDINGS]]),
    }));
    console.log("[LIBRARY] canonical cache bucket resolved", { canonical: lookup.name, available });
  }
  return lookup;
}

async function persistExtractedText(env: Env, fileId: string, text: string) {
  const normalized = normalizeExtractedText(text).slice(0, MAX_OCR_TEXT_LENGTH);
  const { bucket, name } = getExtractionBucket(env);
  const extractedKey = buildExtractedKey(fileId);
  await bucket.put(extractedKey, normalized, {
    httpMetadata: { contentType: "text/plain; charset=utf-8" },
  });
  return { extractedKey, bucket: name, preview: normalized.slice(0, ONE_SHOT_PREVIEW_LIMIT) };
}

async function loadExtractedText(env: Env, { fileId, extractedKey }: { fileId?: string; extractedKey?: string; }) {
  const key = extractedKey?.trim() || (fileId ? buildExtractedKey(fileId) : "");
  if (!key) return null;
  try {
    const { bucket } = getExtractionBucket(env);
    const object = await bucket.get(key);
    if (!object || !object.body) return null;
    const text = await object.text();
    const normalized = normalizeExtractedText(text);
    if (!normalized) return null;
    return { text: normalized.slice(0, MAX_OCR_TEXT_LENGTH), key };
  } catch (err) {
    console.warn("[ASK-FILE] Failed to load extracted text", { key, error: err instanceof Error ? err.message : String(err) });
    return null;
  }
}

async function extractPdfForAsk(env: Env, fileId: string, filename: string): Promise<string> {
  const prompt = [
    `You are a document ingestion (text extraction + OCR) service helping OWEN answer questions.`,
    `Open the attached file (${filename || "Document"}). The file may be text-based or scanned.`,
    `Extract full text in reading order. Preserve headings, bullet points, and table cell contents.`,
    `If the document has pages or slides, prepend each section with markers like [Page 1] or [Slide 2].`,
    `Do NOT summarize. Do NOT fabricate text. Use ? for uncertain characters.`,
  ].join(" ");

  const payload = {
    model: resolveModelId("gpt-4o", env),
    input: [
      {
        role: "user" as const,
        content: [
          { type: "input_text" as const, text: prompt },
          { type: "input_file" as const, file_id: fileId },
        ],
      },
    ],
    max_output_tokens: OCR_MAX_OUTPUT_TOKENS,
  };

  const base = env.OPENAI_API_BASE?.replace(/\/$/, "") || "https://api.openai.com/v1";
  const resp = await retryOpenAI(
    () =>
      fetch(`${base}/responses`, {
        method: "POST",
        headers: {
          authorization: `Bearer ${env.OPENAI_API_KEY}`,
          "content-type": "application/json",
        },
        body: JSON.stringify(payload),
      }),
    `ask-file:extract:${filename}`,
  );

  const data = await safeJson(resp);
  if (!resp.ok) {
    console.error("[ASK-FILE] OpenAI error", data);
    const msg = data?.error?.message || resp.statusText || "PDF extraction failed.";
    throw new Error(msg);
  }
  return extractOutputText(data).trim().slice(0, MAX_EXTRACT_CHARS);
}

async function handlePdfIngest(req: Request, env: Env): Promise<Response> {
  const contentType = req.headers.get("content-type") || "";
  if (!contentType.includes("multipart/form-data")) {
    return json({ error: "Send multipart/form-data with fields: file, fileHash." }, 400);
  }
  const form = await req.formData();
  const file = form.get("file");
  const fileHash = typeof form.get("fileHash") === "string" ? (form.get("fileHash") as string).trim() : "";
  const requestedMaxPages = Number(form.get("maxPages"));
  const maxPages = Number.isFinite(requestedMaxPages) && requestedMaxPages > 0 ? Math.min(requestedMaxPages, DEFAULT_OCR_MAX_PAGES) : DEFAULT_OCR_MAX_PAGES;

  if (!(file instanceof File)) return json({ error: "Missing file upload." }, 400);
  if (!fileHash) return json({ error: "Missing fileHash." }, 400);
  const filename = sanitizeFilename(file.name || "upload.pdf");
  const mimeType = file.type || "application/pdf";
  if (!isLikelyPdfFilename(filename) && !isPdfMimeType(mimeType)) {
    return json({ error: "Only PDF ingestion is supported here." }, 400);
  }

  const { bucket } = getExtractionBucket(env);
  const cache = await loadCachedExtraction(bucket, fileHash);
  if (cache.text) {
    console.log("[PDF-INGEST] cache hit", { fileHash, method: cache.manifest?.method || "cache" });
    return json({
      extractionStatus: "ok",
      method: "cache",
      extractedKey: cache.extractedKey,
      preview: cache.text.slice(0, ONE_SHOT_PREVIEW_LIMIT),
      pageCount: cache.manifest?.pageCount,
      manifest: cache.manifest,
    });
  }
  if (cache.manifest && !cache.text) {
    console.log("[PDF-INGEST] cached manifest without text, requiring OCR", {
      fileHash,
      pageCount: cache.manifest.pageCount,
    });
    return json({
      extractionStatus: "needs_ocr_images",
      method: "ocr",
      fileHash,
      extractedKey: cache.extractedKey,
      pageCount: cache.manifest.pageCount,
      maxPages,
      manifest: cache.manifest,
    });
  }

  const bytes = new Uint8Array(await file.arrayBuffer());
  const extraction = await extractEmbeddedPagesFromPdf(bytes, { sampleCount: PDF_SAMPLE_PAGE_TARGET, maxPages: MAX_EXTRACT_PAGES, allowEarlyStop: true });
  if (extraction.scanned || !extraction.pages.length) {
    const manifest: PdfManifest = {
      fileHash,
      filename,
      method: "ocr",
      pagesProcessed: 0,
      pageCount: extraction.pageCount,
      createdAt: new Date().toISOString(),
    };
    await writeManifest(bucket, manifest);
    console.log("[PDF-INGEST] scanned PDF detected, browser OCR required", {
      fileHash,
      pageCount: extraction.pageCount,
      sampledPagesWithText: extraction.sampledPagesWithText,
    });
    return json({
      extractionStatus: "needs_ocr_images",
      method: "ocr",
      fileHash,
      extractedKey: cache.extractedKey,
      pageCount: extraction.pageCount,
      maxPages,
      manifest,
    });
  }

  const normalized = normalizePages(extraction.pages);
  const finalText = normalized.text.slice(0, MAX_OCR_TEXT_LENGTH);
  const preview = finalText.slice(0, ONE_SHOT_PREVIEW_LIMIT);
  await bucket.put(cache.extractedKey, finalText, {
    httpMetadata: { contentType: "text/plain; charset=utf-8" },
  });
  const manifest: PdfManifest = {
    fileHash,
    filename,
    method: "embedded",
    pagesProcessed: normalized.pagesProcessed,
    pageCount: extraction.pageCount,
    createdAt: new Date().toISOString(),
    preview,
  };
  await writeManifest(bucket, manifest);
  console.log("[PDF-INGEST] embedded text stored", {
    fileHash,
    method: "embedded",
    length: finalText.length,
    pageCount: extraction.pageCount,
  });

  return json({
    extractionStatus: "ok",
    method: "embedded",
    extractedKey: cache.extractedKey,
    preview,
    pageCount: extraction.pageCount,
    manifest,
  });
}

async function handleAskFile(req: Request, env: Env): Promise<Response> {
  const contentType = req.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    const body = await req.json().catch(() => null);
    if (!body || typeof body !== "object" || Array.isArray(body)) {
      return jsonNoStore({ error: "Send JSON { fileId, message, extractedKey? }." }, 400);
    }
    const message = typeof (body as any).message === "string" ? (body as any).message.trim() : "";
    const fileId = typeof (body as any).fileId === "string" ? (body as any).fileId.trim() : "";
    const extractedKey = typeof (body as any).extractedKey === "string" ? (body as any).extractedKey.trim() : "";
    if (!message) return jsonNoStore({ error: "Missing 'message'." }, 400);
    if (!fileId && !extractedKey) return jsonNoStore({ error: "Missing fileId or extractedKey." }, 400);

    const loaded = await loadExtractedText(env, { fileId, extractedKey });
    if (!loaded) {
      return jsonNoStore({ error: "extracted_text_not_found", details: "Extraction not ready or missing for this fileId." }, 404);
    }
    const preview = loaded.text.slice(0, ONE_SHOT_PREVIEW_LIMIT);
    try {
      const answerResult = await answerWithContext(env, loaded.text, message);
      const answer = answerResult.truncated ? appendTruncationNotice(answerResult.fullText) : answerResult.fullText;
      return jsonNoStore({
        answer,
        extractionStatus: "ok",
        method: "stored",
        extractedKey: loaded.key,
        extractedTextPreview: preview,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error("[ASK-FILE] Answering from stored text failed", err);
      return jsonNoStore({ error: "answer_failed", details: msg }, 502);
    }
  }

  if (!contentType.includes("multipart/form-data")) {
    return jsonNoStore({ error: "Send multipart/form-data with fields: message, file" }, 400);
  }
  const form = await req.formData();
  const message = typeof form.get("message") === "string" ? String(form.get("message")).trim() : "";
  const file = form.get("file");
  if (!message) return json({ error: "Missing 'message'." }, 400);
  if (!(file instanceof File)) return json({ error: "Missing 'file'." }, 400);

  const filename = file.name || "upload.bin";
  const mimeType = file.type || "application/octet-stream";
  const bytes = new Uint8Array(await file.arrayBuffer());
  const startTime = Date.now();
  const fileId = generateFileId(filename);
  console.log("[ASK-FILE] start", { filename, mimeType, size: bytes.length, fileId });

  const isPdf = isLikelyPdfFilename(filename) || isPdfMimeType(mimeType);
  const isImage = isLikelyImageFilename(filename) || isImageMimeType(mimeType);
  const preparedImage = isImage ? prepareOcrImageInput(bytes, mimeType || "image/png") : null;

  let extracted = "";
  let extractionStatus: "ok" | "needs_ocr_images" | "error" = "ok";
  let method: "embedded" | "ocr" | "original" = "embedded";

  if (isPdf) {
    try {
      let embedded = "";
      if (PDFJS_AVAILABLE) {
        embedded = await extractEmbeddedTextFromPdf(bytes);
      } else {
        console.log("[ASK-FILE] PDF.js unavailable in Worker; skipping embedded text path.");
      }
      embedded = normalizeExtractedText(embedded);
      if (embedded && embedded.length >= MIN_EMBEDDED_PDF_CHARS) {
        extracted = embedded.slice(0, MAX_EXTRACT_CHARS);
        method = "embedded";
        console.log("[ASK-FILE] Embedded PDF text extracted", { length: extracted.length });
      } else {
        extractionStatus = "needs_ocr_images";
        method = "ocr";
        console.warn("[ASK-FILE] Embedded text too short, requesting browser OCR", {
          length: embedded.length,
          threshold: MIN_EMBEDDED_PDF_CHARS,
        });
      }
    } catch (err) {
      extractionStatus = "needs_ocr_images";
      method = "ocr";
      console.warn("[ASK-FILE] Embedded extraction failed; falling back to OCR images", {
        error: err instanceof Error ? err.message : String(err),
      });
    }
  } else {
    try {
      const ingest = await ingestBinaryForTranscript(env, { bytes, filename, mimeType, sourceKey: filename });
      if (ingest?.text) {
        extracted = ingest.text.trim();
        method = ingest.source === "ocr" ? "ocr" : "embedded";
      }
      if (!extracted && isImage) {
        const imageForOcr = preparedImage || prepareOcrImageInput(bytes, mimeType || "image/png");
        if (!imageForOcr.ok) {
          extractionStatus = "error";
          method = "ocr";
          console.warn("[ASK-FILE] Image bytes failed validation for OCR", {
            filename,
            head: imageForOcr.head,
            mimeType: imageForOcr.mimeType,
          });
          return jsonNoStore(invalidImagePayload({ extractionStatus, method, fileId }), 422);
        }
        try {
          extracted = await requestVisionOcrFromImages(
            env,
            [
              {
                label: "Image",
                dataUrl: imageForOcr.dataUrl,
                sourceKey: filename,
                mimeType: imageForOcr.mimeType,
                signature: imageForOcr.signature,
                byteLength: imageForOcr.byteLength,
                head: imageForOcr.head,
              },
            ],
            filename,
          );
          method = "ocr";
        } catch (err) {
          console.warn("[ASK-FILE] Image OCR failed", err);
        }
        if (!extracted) {
          try {
            const visionFileId = await uploadBytesToOpenAI(env, bytes, filename, "vision");
            if (visionFileId) {
              console.log("[ASK-FILE] uploaded image for OCR", { file_id: visionFileId });
              extracted = await requestVisionTextExtraction(env, visionFileId, filename);
              method = "ocr";
            }
          } catch (err) {
            console.warn("[ASK-FILE] Vision fallback failed", err);
          }
        }
      }
    } catch (err) {
      extractionStatus = "error";
      console.error("[ASK-FILE] Extraction error", err);
      return jsonNoStore(
        {
          extractionStatus,
          extractionError: err instanceof Error ? err.message : "Extraction failed.",
          extractedTextPreview: "",
          answer: "",
        },
        422,
      );
    }
  }

  if (extractionStatus === "needs_ocr_images") {
    return jsonNoStore({
      extractionStatus,
      method,
      fileId,
      pageCap: CLIENT_OCR_PAGE_CAP,
      message: "Browser OCR required; render pages to images and retry.",
    });
  }

  const refusalPhrases = [
    "unable to extract",
    "unable to extract text",
    "unable to read the pdf",
    "can't read the pdf",
    "can't read",
    "cannot read the pdf",
    "cannot read",
    "cannot process documents",
    "cannot access the pdf",
    "don't have the document",
    "process documents directly",
    "paste the text",
    "provide the text",
  ];
  const normalized = normalizeExtractedText(extracted).slice(0, MAX_EXTRACT_CHARS);
  const lower = normalized.toLowerCase();
  const isRefusal = refusalPhrases.some(p => lower.includes(p));
  const tooShort = normalized.length < 50;
  if (!normalized || isRefusal || tooShort) {
    console.warn("[ASK-FILE] extraction invalid, aborting answer", {
      length: normalized.length,
      isRefusal,
      tooShort,
      method,
    });
    return jsonNoStore(
      {
        extractionStatus: "error",
        extractionError: "Extraction returned no usable text. Try OCR images.",
        extractedTextPreview: "",
        answer: "",
      },
      422,
    );
  }

  let persisted: { extractedKey: string; bucket: string; preview: string };
  try {
    persisted = await persistExtractedText(env, fileId, normalized);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[ASK-FILE] Failed to persist extracted text", err);
    return jsonNoStore({ error: "persist_failed", details: msg }, 500);
  }

  try {
    console.log("[ASK-FILE] extraction valid, answering");
    const tableIntent = isTableIntent(message);
    const excerptsAllowed = tableIntent ? false : allowExcerpts(message);
    const tableHeaders = tableIntent ? deriveTableHeaders(message) : undefined;
    debugFormatLog("ask-file format route", {
      tableIntent,
      tableHeaders,
      allowExcerpts: excerptsAllowed,
      textLength: normalized.length,
    });
    const rawAnswerResult = await answerWithContextWithOpts(env, normalized, message, {
      tableOnly: tableIntent,
      tableSchema: tableIntent ? undefined : undefined,
      tableHeaders,
      allowExcerpts: excerptsAllowed,
    });
    const rawAnswer = rawAnswerResult.fullText;
    const preferTable = tableIntent || isInherentlyTabularQuestion(message);
    const formatted = preferTable ? trimAfterFirstTable(rawAnswer) : rawAnswer;
    const repaired = preferTable ? ensureTablePresence(formatted, tableHeaders || deriveTableHeaders(message)) : formatted;
    const sanitized = sanitizeAnswer(repaired, { allowExcerpts: excerptsAllowed });
    const answerCore = finalizeAssistantText(rawAnswer, sanitized, persisted.preview);
    const answer = rawAnswerResult.truncated ? appendTruncationNotice(answerCore) : answerCore;
    debugFormatLog("ask-file format result", {
      preferTable,
      rawLength: rawAnswer.length,
      formattedLength: formatted.length,
      sanitizedLength: sanitized.length,
      finalLength: answer.length,
    });
    return jsonNoStore({
      answer,
      extractedTextPreview: persisted.preview,
      extractionStatus,
      method,
      extractedKey: persisted.extractedKey,
      fileId,
      elapsed_ms: Date.now() - startTime,
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[ASK-FILE] Answering failed", err);
    return jsonNoStore(
      {
        answer: "",
        extractedTextPreview: persisted.preview,
        extractionStatus: "error",
        extractionError: msg,
        method,
        fileId,
        elapsed_ms: Date.now() - startTime,
      },
      502,
    );
  }
}

type GeneratedTextResult = {
  text: string;
  fullText: string;
  truncated: boolean;
  attempts: number;
  finishReason?: string;
  status?: string;
};

async function answerWithContext(env: Env, context: string, question: string): Promise<GeneratedTextResult> {
  return answerWithContextWithOpts(env, context, question, {});
}

async function answerWithContextWithOpts(
  env: Env,
  context: string,
  question: string,
  opts: RetrievedAnswerOptions,
): Promise<GeneratedTextResult> {
  const cleanedContext = cleanRetrievedChunkText(context);
  const system = opts.tableOnly
    ? [
        "You are OWEN. Return ONLY the requested table (plus at most two short sentences of intro).",
        "Do not add headings, extra sections, bullets, or slide text. Do not mention chunks or pages.",
        "Do NOT output internal chunk IDs (C#) or page markers.",
        "Answer in complete sentences and finish cleanly—no dangling fragments.",
        "If a value is not stated in the provided context, write 'Not stated in lecture'.",
        "Table columns: Condition / Disease | Defining labs | Key treatment | Notes (1 line max).",
      ].join(" ")
    : [
        "You are OWEN answering questions about a provided document.",
        "Answer using ONLY the provided document context.",
        "If the context is insufficient, say you don't know.",
        "Answer the user's question directly. Do NOT output slide dumps, page markers, chunk ids, or unrelated headings.",
        "Do NOT output internal chunk IDs (C#) or page markers to the user.",
        "Do NOT copy raw slide text unless explicitly requested.",
        "Answer in complete sentences and end the response cleanly—no dangling words.",
        "If you cite context, use 'Slide X – Title' or 'Page X' style references instead of chunk ids.",
        "Only include information relevant to the question. Stop after the answer; do not append extra sections.",
        opts.allowExcerpts
          ? "If quoting slides is requested, include brief verbatim snippets with citations."
          : "Do not include raw excerpt blocks or a 'From the lecture excerpts' section unless explicitly requested.",
      ].join(" ");

  const userBase: ResponsesInputContent[] = [
    { type: "input_text", text: `Document context:\n${cleanedContext}` },
    { type: "input_text", text: `Question:\n${question}` },
  ];

  const result = await runResponsesWithContinuation(env, {
    label: "ask-context",
    maxOutputTokens: opts.maxOutputTokens ?? ANSWER_MAX_OUTPUT_TOKENS,
    buildPayload: ({ attempt, accumulatedText }) => {
      const continuationParts: ResponsesInputContent[] = [];
      const isContinuation = attempt > 1 || Boolean(accumulatedText);
      if (isContinuation && accumulatedText) {
        continuationParts.push({
          type: "input_text",
          text: `Partial answer so far:\n${clipAnswerTail(accumulatedText)}`,
        });
      }
      continuationParts.push({
        type: "input_text",
        text: "Continue exactly where you left off. Do not repeat earlier lines or headings. Do not restart tables. Preserve the same narrative formatting.",
      });

      return {
        model: resolveModelId("gpt-5-mini", env),
        max_output_tokens: opts.maxOutputTokens ?? ANSWER_MAX_OUTPUT_TOKENS,
        input: [
          { role: "system" as const, content: [{ type: "input_text" as const, text: system }] },
          { role: "user" as const, content: [...userBase, ...continuationParts] },
        ],
      };
    },
  });

  return result;
}

type ResponseCallResult = {
  text: string;
  finishReason?: string;
  status?: string;
  outputTokens?: number;
  incompleteReason?: string;
};

async function runResponsesWithContinuation(
  env: Env,
  opts: {
    buildPayload: (state: { attempt: number; accumulatedText: string }) => Record<string, unknown>;
    label: string;
    maxAttempts?: number;
    maxOutputTokens?: number;
    initialText?: string;
    onSegment?: (segment: string, state: { attempt: number; accumulatedText: string }) => Promise<void> | void;
  },
): Promise<GeneratedTextResult> {
  const attemptsLimit = Math.max(1, opts.maxAttempts ?? ANSWER_MAX_CONTINUATIONS + 1);
  const maxTokens = Math.min(ANSWER_MAX_OUTPUT_TOKENS, Math.max(800, opts.maxOutputTokens ?? ANSWER_MAX_OUTPUT_TOKENS));
  let accumulated = opts.initialText || "";
  const baseLength = accumulated.length;
  let attempts = 0;
  let truncated = false;
  let lastFinish: string | undefined;
  let lastStatus: string | undefined;
  let continuationPrefixed = false;

  while (attempts < attemptsLimit) {
    attempts += 1;
    const payload = opts.buildPayload({ attempt: attempts, accumulatedText: accumulated }) || {};
    const resolvedMaxOutputTokens = (payload as any)?.max_output_tokens ?? maxTokens;
    const payloadWithMax = {
      ...payload,
      max_output_tokens: resolvedMaxOutputTokens,
    };
    const result = await callResponsesOnce(env, payloadWithMax, `${opts.label || "ask"}#${attempts}`);
    lastFinish = result.finishReason || lastFinish;
    lastStatus = result.status || lastStatus;
    let segment = (result.text || "").trim();
    if (segment) {
      if (attempts > 1 || accumulated) {
        const prepared = prepareContinuationSegment(segment, accumulated, continuationPrefixed);
        segment = prepared.text;
        continuationPrefixed = prepared.prefixed;
      }
      accumulated = combineSegments([accumulated, segment]);
      if (opts.onSegment) {
        await opts.onSegment(segment, { attempt: attempts, accumulatedText: accumulated });
      }
    }
    const needsMore = shouldContinueAnswer(result, maxTokens, segment);
    if (!needsMore) {
      truncated = false;
      break;
    }
    if (attempts >= attemptsLimit) {
      truncated = true;
    }
  }

  const fullText = accumulated;
  const text = fullText.slice(baseLength);

  return {
    text,
    fullText,
    truncated,
    attempts,
    finishReason: lastFinish,
    status: lastStatus,
  };
}

async function callResponsesOnce(env: Env, payload: Record<string, unknown>, label: string): Promise<ResponseCallResult> {
  const base = env.OPENAI_API_BASE?.replace(/\/$/, "") || "https://api.openai.com/v1";
  const resp = await retryOpenAI(
    () =>
      fetch(`${base}/responses`, {
        method: "POST",
        headers: {
          authorization: `Bearer ${env.OPENAI_API_KEY}`,
          "content-type": "application/json",
          "OpenAI-Beta": "assistants=v2",
        },
        body: JSON.stringify(payload),
      }),
    label,
  );

  const data = await safeJson(resp);
  if (!resp.ok) {
    const msg = data?.error?.message || resp.statusText || "Answering failed.";
    throw new Error(msg);
  }
  return {
    text: extractOutputText(data).trim(),
    finishReason: extractFinishReason(data),
    status: extractResponseStatus(data),
    outputTokens: extractOutputTokens(data),
    incompleteReason: extractIncompleteReason(data),
  };
}

async function runChatCompletionsWithContinuation(
  env: Env,
  opts: {
    buildMessages: (state: { attempt: number; accumulatedText: string }) => ChatMessage[];
    model: string;
    label: string;
    maxAttempts?: number;
    maxOutputTokens?: number;
    onSegment?: (segment: string, state: { attempt: number; accumulatedText: string }) => Promise<void> | void;
  },
): Promise<GeneratedTextResult> {
  const attemptsLimit = Math.max(1, opts.maxAttempts ?? ANSWER_MAX_CONTINUATIONS + 1);
  const maxTokens = Math.min(ANSWER_MAX_OUTPUT_TOKENS, Math.max(800, opts.maxOutputTokens ?? ANSWER_MAX_OUTPUT_TOKENS));
  let accumulated = "";
  let attempts = 0;
  let truncated = false;
  let lastFinish: string | undefined;

  while (attempts < attemptsLimit) {
    attempts += 1;
    const messages = opts.buildMessages({ attempt: attempts, accumulatedText: accumulated }) || [];
    const result = await callChatCompletionOnce(env, {
      messages,
      model: opts.model,
      maxTokens,
      label: `${opts.label}#${attempts}`,
    });
    lastFinish = result.finishReason || lastFinish;
    const segment = (result.text || "").trim();
    if (segment) {
      accumulated = combineSegments([accumulated, segment]);
      if (opts.onSegment) {
        await opts.onSegment(segment, { attempt: attempts, accumulatedText: accumulated });
      }
    }
    const needsMore = shouldContinueAnswer(result, maxTokens, segment);
    if (!needsMore) {
      truncated = false;
      break;
    }
    if (attempts >= attemptsLimit) {
      truncated = true;
    }
  }

  return {
    text: accumulated,
    fullText: accumulated,
    truncated,
    attempts,
    finishReason: lastFinish,
  };
}

async function callChatCompletionOnce(
  env: Env,
  input: { messages: ChatMessage[]; model: string; maxTokens: number; label: string },
): Promise<ResponseCallResult> {
  const base = env.OPENAI_API_BASE?.replace(/\/$/, "") || "https://api.openai.com/v1";
  const resp = await retryOpenAI(
    () =>
      fetch(`${base}/chat/completions`, {
        method: "POST",
        headers: {
          authorization: `Bearer ${env.OPENAI_API_KEY}`,
          "content-type": "application/json",
        },
        body: JSON.stringify({
          model: input.model,
          messages: input.messages,
          max_tokens: input.maxTokens,
          max_output_tokens: input.maxTokens,
          stream: false,
        }),
      }),
    input.label,
  );
  const data = await safeJson(resp);
  if (!resp.ok) {
    const msg = data?.error?.message || resp.statusText || "Chat completion failed.";
    throw new Error(msg);
  }
  return {
    text: extractChatCompletionContent(data).trim(),
    finishReason: data?.choices?.[0]?.finish_reason,
    status: resp.status === 200 ? "completed" : undefined,
    outputTokens: data?.usage?.completion_tokens ?? data?.usage?.output_tokens,
  };
}

function extractFinishReason(payload: any): string | undefined {
  return (
    payload?.response?.output?.[0]?.finish_reason ||
    payload?.output?.[0]?.finish_reason ||
    payload?.response?.stop_reason ||
    payload?.stop_reason ||
    payload?.response?.incomplete_details?.reason ||
    payload?.incomplete_details?.reason ||
    undefined
  );
}

function extractResponseStatus(payload: any): string | undefined {
  return payload?.response?.status || payload?.status || undefined;
}

function extractOutputTokens(payload: any): number | undefined {
  const candidates = [
    payload?.usage?.output_tokens,
    payload?.usage?.completion_tokens,
    payload?.response?.usage?.output_tokens,
    payload?.response?.usage?.completion_tokens,
  ];
  for (const value of candidates) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }
  }
  return undefined;
}

function extractIncompleteReason(payload: any): string | undefined {
  return payload?.incomplete_details?.reason || payload?.response?.incomplete_details?.reason || undefined;
}

function shouldContinueAnswer(result: ResponseCallResult, maxOutputTokens: number, segmentText: string): boolean {
  const finish = (result.finishReason || "").toLowerCase();
  const incomplete = (result.incompleteReason || "").toLowerCase();
  if (finish === "length" || finish === "max_tokens") return true;
  if (incomplete.includes("max") || incomplete.includes("length")) return true;
  if (result.status && result.status !== "completed" && result.status !== "finished") return true;
  const nearCap = typeof result.outputTokens === "number" && result.outputTokens >= maxOutputTokens - 20;
  if (nearCap) return true;
  if (looksTruncatedText(segmentText) && (nearCap || segmentText.length >= 1200)) return true;
  return false;
}

function looksTruncatedText(text: string): boolean {
  const trimmed = (text || "").trim();
  if (!trimmed) return false;
  if (trimmed.endsWith("</table>")) return false;
  if (/[.!?…)]$/.test(trimmed)) return false;
  if (trimmed.endsWith("...")) return true;
  const lastLine = trimmed.split(/\r?\n/).pop()?.trim() || "";
  if (/[,;:/-]\s*$/.test(lastLine)) return true;
  if (trimmed.length > 500 && !/[.!?…)]$/.test(lastLine)) return true;
  return false;
}

function combineSegments(segments: string[]): string {
  return segments.reduce((acc, segment) => {
    if (!segment) return acc;
    if (!acc) return segment;
    const needsSpace =
      !acc.endsWith("\n") &&
      !segment.startsWith("\n") &&
      !acc.endsWith(" ") &&
      !segment.startsWith(" ");
    return acc + (needsSpace ? " " : "") + segment;
  }, "");
}

function normalizeHeadingText(line: string): string {
  return line
    .replace(/^#{1,6}\s*/, "")
    .replace(/^[-*•]\s*/, "")
    .replace(/\*\*/g, "")
    .replace(/[:.]+$/, "")
    .trim()
    .toLowerCase();
}

function isLikelyHeading(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed) return false;
  if (/^#{1,6}\s+\S/.test(trimmed)) return true;
  if (/^\*\*[A-Za-z0-9].*\*\*:?$/.test(trimmed)) return true;
  if (/^[A-Z][A-Za-z0-9 ()./'-]{1,80}:?$/.test(trimmed) && !/^[-*•]/.test(trimmed)) {
    const words = trimmed.split(/\s+/);
    if (words.length <= 10) return true;
  }
  return false;
}

function findTrailingHeading(text: string): string | null {
  const lines = (text || "").split(/\r?\n/).slice(-8).reverse();
  for (const line of lines) {
    if (isLikelyHeading(line)) {
      return normalizeHeadingText(line);
    }
  }
  return null;
}

function stripDuplicateLeadingHeaders(segment: string, prior: string): string {
  const priorHeading = findTrailingHeading(prior);
  const lines = (segment || "").split(/\r?\n/);
  while (lines.length) {
    const line = lines[0];
    if (!isLikelyHeading(line)) break;
    if (priorHeading && normalizeHeadingText(line) === priorHeading) {
      lines.shift();
      continue;
    }
    break;
  }
  while (lines.length && !lines[0].trim()) lines.shift();
  return lines.join("\n");
}

function prepareContinuationSegment(segment: string, prior: string, alreadyPrefixed: boolean): { text: string; prefixed: boolean } {
  let cleaned = stripDuplicateLeadingHeaders(segment, prior);
  let prefixed = alreadyPrefixed;
  if (cleaned.trim()) {
    const hasPrefix = /^\(continuing/i.test(cleaned.trimStart());
    if (!prefixed && !hasPrefix) {
      cleaned = `(Continuing…)\n${cleaned.trimStart()}`;
      prefixed = true;
    }
  }
  return { text: cleaned, prefixed };
}

function clipAnswerTail(answer: string, limit = CONTINUATION_TAIL_CHARS): string {
  const trimmed = (answer || "").trim();
  if (trimmed.length <= limit) return trimmed;
  return trimmed.slice(-limit);
}

function appendTruncationNotice(answer: string): string {
  const trimmed = (answer || "").trim();
  if (!trimmed) return TRUNCATION_NOTICE;
  if (trimmed.includes(TRUNCATION_NOTICE)) return trimmed;
  const separator = /\n\s*$/.test(trimmed) ? "" : (/[.!?]$/.test(trimmed) ? " " : "\n\n");
  return `${trimmed}${separator}${TRUNCATION_NOTICE}`;
}

type LibraryIntent = {
  isBroad: boolean;
  wantsTable: boolean;
  wantsConditionLabsTable: boolean;
  listIntent: boolean;
};

function isTableIntent(prompt: string): boolean {
  const lower = (prompt || "").toLowerCase();
  if (!lower) return false;
  if (/\b(table|tabulate|columns|make a table|create a table|put in a table|table of)\b/.test(lower)) return true;
  return false;
}

function isInherentlyTabularQuestion(prompt: string): boolean {
  const lower = (prompt || "").toLowerCase();
  if (!lower) return false;
  const phrases = ["reference range", "ref range", "cut-off", "cut off"];
  if (phrases.some(phrase => lower.includes(phrase))) return true;
  const keywords = [
    "lab",
    "labs",
    "laboratory",
    "threshold",
    "cutoff",
    "dose",
    "dosing",
    "dosage",
    "sensitivity",
    "specificity",
    "value",
    "values",
  ];
  return keywords.some(keyword => lower.includes(keyword));
}

function allowExcerpts(prompt: string): boolean {
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

function looksLikeAttributeList(answer: string): boolean {
  if (!answer) return false;
  const lines = (answer || "").split(/\r?\n/);
  let hits = 0;
  lines.forEach(line => {
    const trimmed = line.trim();
    if (!/^[-*•]/.test(trimmed)) return;
    const body = trimmed.replace(/^[-*•]\s*/, "");
    if (/^([A-Za-z0-9 ()./+%~-]{2,})\s*[:–-]\s+.+/.test(body)) {
      hits += 1;
    }
  });
  return hits >= 3;
}

function deriveTableHeaders(prompt: string): string[] {
  const lower = (prompt || "").toLowerCase();
  if (/\bdisease/.test(lower) && /\bgenetic/.test(lower)) {
    return ["Disease", "Defining features", "Genetic marker(s) (if relevant)"];
  }
  if (/\bdisease/.test(lower) && /\bfeature/.test(lower)) {
    return ["Disease", "Key features", "Notes"];
  }
  if (/\bdrug/.test(lower) && /\bdose|dosing|dosage/.test(lower)) {
    return ["Drug", "Indication", "Dose", "Notes"];
  }
  return ["Item", "Details"];
}

const LIBRARY_HEADER_KEYWORDS = [
  "WORKUP",
  "CAUSES",
  "ETIOLOGY",
  "ETIOLOGIES",
  "DIFFERENTIAL",
  "DIAGNOSIS",
  "TABLE",
  "ALGORITHM",
  "APPROACH",
  "SUMMARY",
  "OVERVIEW",
];

function detectLibraryIntent(question: string): LibraryIntent {
  const lower = (question || "").toLowerCase();
  const listIntent = /\b(list|all|table|catalog|enumerat|differential|compare|versus|vs|outline|overview|summary)\b/i.test(
    lower,
  );
  const broadTriggers = [
    "table of",
    "list all",
    "list the",
    "all causes",
    "all etiolog",
    "all of the",
    "compare",
    "versus",
    "vs",
    "algorithm",
    "workup",
    "approach",
    "differential",
    "overview",
    "summary",
    "outline",
    "causes",
    "etiology",
    "etiologies",
    "etiologic",
  ];
  const isBroad = broadTriggers.some(trigger => lower.includes(trigger)) || /\btable\b/.test(lower);
  const wantsTable = isTableIntent(question);
  const wantsConditionLabsTable =
    wantsTable &&
    (/\blab\b|\blabs\b|laboratory|electrolyte|osm|sodium|potassium|urine|serum/.test(lower) ||
      /\bvalues?\b/.test(lower));
  return { isBroad, wantsTable, wantsConditionLabsTable, listIntent };
}

function isHeaderChunk(text: string): boolean {
  if (!text) return false;
  const upper = text.toUpperCase();
  const headerHit = LIBRARY_HEADER_KEYWORDS.some(keyword => upper.includes(keyword));
  const uppercaseWords = text.match(/\b[A-Z]{3,}\b/g) || [];
  return headerHit || uppercaseWords.length >= 6;
}

function isTableLikeChunk(text: string): boolean {
  if (!text) return false;
  const pipeCount = (text.match(/\|/g) || []).length;
  const tabCount = (text.match(/\t/g) || []).length;
  const multiSpaceLines = (text.match(/\n[^\n]{0,80} {3,}[^\n]{0,80}/g) || []).length;
  return pipeCount >= 2 || tabCount >= 2 || multiSpaceLines >= 3;
}

function hasDenseAcronyms(text: string): boolean {
  if (!text) return false;
  const acronyms = text.match(/\b[A-Z]{2,5}\b/g) || [];
  const unique = new Set(acronyms);
  return unique.size >= 4;
}

function selectLibraryChunks(question: string, chunks: RetrievalChunk[], intent: LibraryIntent): RetrievalChunk[] {
  if (!Array.isArray(chunks) || !chunks.length) return [];

  if (!intent.isBroad) {
    const ranked = rankChunks(question, chunks, RETRIEVAL_TOP_K);
    return ranked.length ? ranked : chunks.slice(0, Math.min(RETRIEVAL_TOP_K, chunks.length));
  }

  const selected = new Map<number, RetrievalChunk>();
  let contextCharCount = 0;
  const addChunk = (chunk: RetrievalChunk) => {
    if (!chunk || selected.has(chunk.index)) return;
    if (contextCharCount + chunk.text.length > LIBRARY_BROAD_MAX_CONTEXT_CHARS) return;
    selected.set(chunk.index, chunk);
    contextCharCount += chunk.text.length;
  };

  const broadTopK = Math.min(chunks.length, LIBRARY_BROAD_TOP_K);
  const ranked = rankChunks(question, chunks, broadTopK);
  const baseCandidates = ranked.length ? ranked : chunks.slice(0, broadTopK);
  baseCandidates.forEach(addChunk);

  const headerCandidates = chunks.filter(chunk => isHeaderChunk(chunk.text)).slice(0, broadTopK);
  headerCandidates.forEach(addChunk);

  const tableCandidates = chunks.filter(chunk => isTableLikeChunk(chunk.text)).slice(0, broadTopK);
  tableCandidates.forEach(addChunk);

  const acronymCandidates = chunks.filter(chunk => hasDenseAcronyms(chunk.text)).slice(0, broadTopK);
  acronymCandidates.forEach(addChunk);

  const needsCoverage =
    selected.size < LIBRARY_BROAD_MIN_CHUNKS || contextCharCount < LIBRARY_BROAD_MIN_CHARS;
  if (needsCoverage && contextCharCount < LIBRARY_BROAD_MAX_CONTEXT_CHARS) {
    for (const chunk of chunks) {
      addChunk(chunk);
      if (
        selected.size >= LIBRARY_BROAD_MIN_CHUNKS &&
        contextCharCount >= LIBRARY_BROAD_MIN_CHARS
      ) {
        break;
      }
    }
  }

  return Array.from(selected.values()).sort((a, b) => a.index - b.index);
}

function trimAfterFirstTable(answer: string): string {
  if (!answer) return answer;
  const htmlClose = answer.indexOf("</table>");
  if (htmlClose !== -1) {
    return answer.slice(0, htmlClose + "</table>".length);
  }

  const lines = answer.split("\n");
  let tableStart = -1;
  let tableEnd = lines.length;
  for (let i = 0; i < lines.length - 1; i++) {
    const line = lines[i];
    const next = lines[i + 1] || "";
    const looksLikeRow = /\|/.test(line) && !/^```/.test(line);
    const looksLikeDivider = /^\s*\|?\s*:?-{2,}/.test(next);
    if (looksLikeRow && looksLikeDivider) {
      tableStart = i;
      break;
    }
  }
  if (tableStart === -1) return answer;
  for (let j = tableStart + 2; j < lines.length; j++) {
    if (!lines[j].trim()) {
      tableEnd = j;
      break;
    }
  }
  return lines.slice(0, tableEnd).join("\n").trim();
}

const SECTION_LABEL_WHITELIST = new Set([
  "response",
  "summary",
  "diagnosis",
  "labs",
  "treatment",
  "warnings",
  "plan",
  "next steps",
]);

function sanitizeAnswer(answer: string, opts: { allowExcerpts?: boolean } = {}): string {
  if (!answer) return answer;
  const { allowExcerpts = false } = opts;
  const cleanedLines: string[] = [];
  const lines = answer.replace(/\[?\bC\d+\b\]?:?/gi, " ").split(/\r?\n/);

  for (const line of lines) {
    const trimmed = line.trim();
    if (!allowExcerpts) {
      if (/from the lecture excerpts:/i.test(trimmed)) break;
      if (/^(?:-|\*|•)?\s*(slide|page)\s*\d+/i.test(trimmed)) break;
    }
    if (!trimmed) {
      if (cleanedLines.length && cleanedLines[cleanedLines.length - 1] !== "") {
        cleanedLines.push("");
      }
      continue;
    }
    if (/^---\s*page\s*\d+\s*---/i.test(trimmed)) continue;
    if (/^c\d+\s*:?\s*-{2,}/i.test(trimmed)) continue;
    if (/^c\d+\s*:?\s*$/i.test(trimmed)) continue;

    const isTiny = trimmed.length <= 3 && !/^[-•*]\s+\S/.test(trimmed);
    const looksCourseHeader =
      /^kcu[-\s]?com/i.test(trimmed) ||
      /^dr\./i.test(trimmed) ||
      (/^[A-Z0-9 .,'/-]+$/.test(trimmed) && trimmed === trimmed.toUpperCase() && trimmed.length > 8);
    const headingMatch = /^([A-Za-z][A-Za-z ]{0,50})[:.]?$/.exec(trimmed);
    const normalizedHeading = headingMatch ? headingMatch[1].trim().toLowerCase() : "";
    const isHeadingOnly = headingMatch && headingMatch[0].length === trimmed.length;
    const allowedHeading = normalizedHeading && SECTION_LABEL_WHITELIST.has(normalizedHeading);

    if (isTiny) continue;
    if (looksCourseHeader) continue;
    if (isHeadingOnly && !allowedHeading) continue;

    cleanedLines.push(line);
  }

  let text = cleanedLines.join("\n");
  text = text
    .replace(/\[?\bC\d+\b\]?:?/gi, "")
    .replace(/^\s*---\s*page\s*\d+\s*---\s*$/gim, "")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]+\n/g, "\n")
    .trim();

  const hasTable = /\|/.test(text) && /\n/.test(text);
  text = text.replace(/\s*(or seiz)\w*$/i, "").trim();

  if (!/[.!?…]\s*$/.test(text) && !hasTable) {
    const lastStop = Math.max(text.lastIndexOf("."), text.lastIndexOf("?"), text.lastIndexOf("!"), text.lastIndexOf("…"));
    if (lastStop !== -1) {
      text = text.slice(0, lastStop + 1).trim();
    }
  }

  if (!/[.!?…]\s*$/.test(text) && !hasTable) {
    const trailingWord = text.match(/[A-Za-z]+$/);
    if (trailingWord) {
      const boundary = Math.max(
        text.lastIndexOf("."),
        text.lastIndexOf("?"),
        text.lastIndexOf("!"),
        text.lastIndexOf("…"),
      );
      if (boundary > -1 && boundary < text.length - 1) {
        text = text.slice(0, boundary + 1).trim();
      } else if (text.split(/\s+/).length > 6) {
        text = text.split(/\s+/).slice(0, -1).join(" ").trim();
      }
    }
  }

  if (text && !/[.!?…]\s*$/.test(text) && !hasTable) {
    text = text.replace(/\s+\n/g, "\n").trimEnd();
    text += ".";
  }

  return text || answer.trim();
}

function sanitizeJunkTail(answer: string, allowExcerpts = false): string {
  return sanitizeAnswer(answer, { allowExcerpts });
}

function cleanRetrievedChunkText(text: string): string {
  const lines = (text || "").split(/\r?\n/);
  const cleaned: string[] = [];
  lines.forEach(line => {
    const trimmed = line.trim();
    if (!trimmed) {
      cleaned.push("");
      return;
    }
    if (/^c\d+:/i.test(trimmed)) return;
    if (/^[A-Za-z]{1,3}$/.test(trimmed)) return;
    if (/^kcu[-\s]?com/i.test(trimmed)) return;
    if (/^dr\./i.test(trimmed) && trimmed === trimmed.toUpperCase()) return;
    if (/^---\s*page\s*\d+\s*---/i.test(trimmed)) return;
    cleaned.push(line);
  });
  const result = cleaned.join("\n").trim();
  return result || text;
}

type ChunkReference = {
  label: string;
  title: string;
  type: "slide" | "page";
  number?: number;
  slide?: number;
  page?: number;
};

function mapChunkToSlideRef(chunk: { index: number; text: string }): ChunkReference {
  const raw = chunk.text || "";
  const lines = raw.split(/\n/).map(l => l.trim()).filter(Boolean);
  const pageMatch = raw.match(/---\s*page\s*(\d+)\s*---/i) || raw.match(/\bpage\s+(\d+)\b/i);
  const slideMatch = raw.match(/\bslide\s+(\d+)\b/i);
  const number = slideMatch ? Number(slideMatch[1]) : pageMatch ? Number(pageMatch[1]) : chunk.index + 1;
  const type: "slide" | "page" = slideMatch ? "slide" : pageMatch ? "page" : "slide";
  const label = `${type === "slide" ? "Slide" : "Page"} ${Number.isFinite(number) ? number : chunk.index + 1}`;
  const title = lines.find(l => l.length > 8 && !/^---\s*page/i.test(l) && !/^slide\s+\d+/i.test(l)) || "";
  const slide = type === "slide" ? (Number.isFinite(number) ? number : undefined) : undefined;
  const page = type === "page" ? (Number.isFinite(number) ? number : undefined) : undefined;
  return { label, title, type, number: Number.isFinite(number) ? number : undefined, slide, page };
}

function replaceChunkMarkers(answer: string, chunks: Array<{ index: number; text: string }>) {
  if (!answer) return answer;
  const chunkMap = new Map<number, ChunkReference>();
  chunks.forEach(chunk => {
    const ref = mapChunkToSlideRef(chunk);
    chunkMap.set(chunk.index + 1, ref);
  });
  return answer.replace(/\[?C(\d+)\]?:?/gi, (_m, numStr) => {
    const num = Number(numStr);
    const ref = chunkMap.get(num);
    if (!ref) return "";
    const title = ref.title ? ` – ${ref.title.slice(0, 80)}` : "";
    return `(${ref.label}${title})`;
  });
}

function buildReferencesFromChunks(chunks: Array<{ index: number; text: string }>, docId?: string) {
  const seen = new Set<string>();
  const refs: Array<{ slide?: number; page?: number; title?: string; docId?: string }> = [];
  chunks.forEach(chunk => {
    const ref = mapChunkToSlideRef(chunk);
    const key = `${ref.type}:${ref.number || ref.label.toLowerCase()}`;
    if (seen.has(key)) return;
    seen.add(key);
    refs.push({
      slide: ref.slide,
      page: ref.page,
      title: ref.title || undefined,
      docId,
    });
  });
  return refs;
}

function buildEvidenceFromChunks(chunks: Array<{ index: number; text: string }>, docId?: string) {
  return chunks
    .map(chunk => {
      const ref = mapChunkToSlideRef(chunk);
      const cleaned = cleanRetrievedChunkText(chunk.text || "")
        .replace(/^\s*---\s*page\s*\d+\s*---\s*$/gim, "")
        .trim();
      const text = cleaned.length > 480 ? `${cleaned.slice(0, 480).trim()}…` : cleaned;
      if (!text) return null;
      return {
        slide: ref.slide,
        page: ref.page,
        title: ref.title || undefined,
        docId,
        excerpt: text,
      };
    })
    .filter((entry): entry is { slide?: number; page?: number; title?: string; excerpt: string } => Boolean(entry));
}

function hasMarkdownTable(answer: string): boolean {
  const lines = (answer || "").split(/\r?\n/);
  for (let i = 0; i < lines.length - 1; i += 1) {
    if (/^\s*\|.+\|\s*$/.test(lines[i]) && /^\s*\|\s*:?-{2,}/.test(lines[i + 1])) {
      return true;
    }
  }
  return false;
}

function buildFallbackTable(headers?: string[]): string {
  const cols = (headers && headers.length ? headers : ["Item", "Details"]).map(h => h.trim()).filter(Boolean);
  if (!cols.length) cols.push("Item", "Details");
  const headerRow = `| ${cols.join(" | ")} |`;
  const divider = `| ${cols.map(() => "---").join(" | ")} |`;
  const emptyRow = `| No data found in lecture | ${cols.length > 1 ? cols.slice(1).map(() => "—").join(" | ") : "—"} |`;
  return [headerRow, divider, emptyRow].join("\n");
}

function ensureTablePresence(answer: string, headers?: string[]): string {
  if (hasMarkdownTable(answer)) return answer.trim();
  const lines = (answer || "")
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(Boolean)
    .map(l => l.replace(/^[-*•]\s*/, ""));

  const rows: string[][] = [];
  const cols = (headers && headers.length ? headers : ["Item", "Details"]).map(h => h.trim()).filter(Boolean);
  if (!cols.length) cols.push("Item", "Details");
  lines.forEach(line => {
    const match = /^([^:–-]+)\s*[:–-]\s*(.+)$/.exec(line);
    if (match) {
      const key = match[1].trim();
      const val = match[2].trim();
      if (key && val) {
        rows.push([key, val]);
      }
    }
  });

  if (!rows.length) {
    return buildFallbackTable(headers);
  }
  const headerRow = `| ${cols.join(" | ")} |`;
  const divider = `| ${cols.map(() => "---").join(" | ")} |`;
  const body = rows
    .slice(0, 30)
    .map(row => {
      const filled = [...row];
      while (filled.length < cols.length) filled.push("—");
      return `| ${filled.slice(0, cols.length).join(" | ")} |`;
    })
    .join("\n");
  const table = [headerRow, divider, body].join("\n");
  if (!hasMarkdownTable(table)) {
    return buildFallbackTable(headers);
  }
  return table;
}

type RetrievedAnswerOptions = {
  mode?: "default" | "library";
  broad?: boolean;
  tableSchema?: "conditionLabs" | "lectureTable";
  tableOnly?: boolean;
  allowExcerpts?: boolean;
  tableHeaders?: string[];
  listIntent?: boolean;
  requestId?: string;
  contextNotes?: string;
  maxOutputTokens?: number;
  initialAnswerText?: string;
};

function buildLibrarySystemPrompt(opts: RetrievedAnswerOptions = {}): string {
  const parts = [
    "You are a medical educator synthesizing lecture material into clinically meaningful explanations.",
    "Treat the provided excerpts as the full lecture; synthesize across all of them without adding outside facts.",
    "Prioritize clear reasoning and human-readable flow. Use explicit language such as 'The lecture states…' or 'Genetic markers are not specified in the lecture' when details are missing.",
    "If a requested detail is missing from the excerpts, state 'not stated in lecture' instead of speculating.",
    "Do not mention retrieval, chunks, or limitations; never expose chunk IDs (C#) or page markers.",
    "Do NOT copy raw slide text unless explicitly requested.",
    "Answer in complete sentences and finish cleanly—no dangling fragments.",
    "Default format: a 1–2 sentence summary followed by clinically grouped bullets (diagnosis/features/management as relevant) written in clinician voice.",
    "Use tables only when the user explicitly asks for one or when the data are inherently tabular (e.g., labs, thresholds, dosing). Otherwise prefer narrative + bullets.",
    "If age distinctions matter, group under Pediatric, Adult, and Not age-specific when helpful; skip headings that are irrelevant and state when the lecture does not specify an age group.",
    "Keep only information relevant to the question and stop after the answer.",
    "Treat any context notes as internal scaffolding—do not quote or paraphrase them; compose the final answer fresh from the lecture content.",
    "If you include structure, use only the following labels when they naturally fit: Response, Summary, Plan, Diagnosis, Labs, Treatment, Warnings, Next steps.",
    "If you use references, present them as 'Slide X – Title' when known, otherwise 'Page X'. Never show chunk ids.",
  ];
  if (!opts.allowExcerpts) {
    parts.push("Do not include excerpt blocks or a 'From the lecture excerpts' section unless the user explicitly asked.");
  }
  if (opts.broad) {
    parts.push(
      "For broad synthesis tasks, compile a comprehensive list across the lecture rather than a narrow sample, but keep the narrative flow.",
    );
  }
  if (opts.listIntent) {
    parts.push(
      "User intent hints at listing/categorizing entities. Use the default summary + clinically grouped bullets, and group Pediatric vs Adult vs Not age-specific only if the lecture distinguishes them.",
      "Tables are optional here—use them only if the content is inherently tabular or the user asked. Otherwise, stay with prose and bullets.",
      "Always note when genetics, age group, or other details are not specified in the lecture.",
      "Do not force headers that do not fit the content.",
    );
  }
  if (opts.tableOnly && Array.isArray(opts.tableHeaders) && opts.tableHeaders.length) {
    const columns = opts.tableHeaders.join(" | ");
    parts.push(
      `Output ONLY one Markdown table (pipe syntax, no code fences). Columns: ${columns}.`,
      "Do not include any extra text before or after the table. If a value is missing, write 'Not stated'.",
    );
  } else if (opts.tableSchema === "lectureTable") {
    parts.push(
      "Output ONLY one Markdown table (pipe syntax, no code fences). You may include at most a single, short introductory sentence before the table.",
      "After the table, output nothing else—no extra sections, slides, or bullets.",
      "Table columns: Condition/Disease | Defining labs | Key treatment | Notes (1 line max).",
      "Use lecture wording; if a lab or treatment is absent, write 'Not stated in lecture'. Do not invent numbers or slide labels.",
    );
  } else if (opts.tableSchema === "conditionLabs") {
    parts.push(
      "Output a single Markdown table (pipe syntax, no code fences) with columns: Condition / Etiology | Volume status | Serum Osm | Urine Osm | Urine Na | Key additional labs | Notes.",
      "Use qualitative descriptors from the lecture (e.g., high/low/inappropriately high); do not invent numeric ranges or reference intervals.",
      "If a value is not specified in the lecture, write 'not stated in lecture'. Keep the table tidy and avoid preambles.",
    );
  }
  parts.push(
    "Match the requested format exactly and keep any notes concise.",
    "When citing the lecture, use slide or page references with short titles, e.g., 'Slide 11 – Diagnosis of Hyperkalemia'.",
    "Do NOT use internal chunk ids like C1, C2, etc., in the final answer.",
  );
  return parts.join(" ");
}

function buildDefaultSystemPrompt(opts: RetrievedAnswerOptions = {}): string {
  if (opts.tableOnly) {
    const columns = Array.isArray(opts.tableHeaders) && opts.tableHeaders.length
      ? opts.tableHeaders
      : ["Condition / Disease", "Defining labs", "Key treatment", "Notes (1 line max)"];
    return [
      "You are OWEN. Return ONLY the requested table (plus at most two short sentences of intro).",
      "Do not add headings, extra sections, bullets, or slide text. Do not mention chunks or pages.",
      "Do NOT output internal chunk IDs (C#) or page markers.",
      "Answer in complete sentences and finish cleanly—no dangling fragments.",
      "If a value is not stated in the provided context, write 'Not stated in lecture'.",
      `Table columns: ${columns.join(" | ")}.`,
      "After the table, output nothing else.",
    ].join(" ");
  }
  return [
    "You are OWEN answering questions grounded in retrieved document chunks.",
    "Use ONLY the provided chunks; do not add outside facts.",
    "If the chunks lack the answer, say you don't know briefly.",
    "Answer the user's question directly. Do NOT output slide dumps, page markers, chunk ids, or unrelated headings.",
    "Do NOT output internal chunk IDs (C#) or page markers to the user.",
    "Do NOT copy raw slide text unless explicitly requested.",
    "Answer in complete sentences and end the response cleanly—no dangling words.",
    "If you include references, format them as 'Slide X – Title' when known, otherwise 'Page X'.",
    "Only include information relevant to the question. Stop after the answer; do not append extra sections.",
    "Organize answers with: Summary (1–2 sentences); Key points (3–6 bullets); Clinical thresholds/steps (bulleted when applicable); Pitfalls / practical notes (1–3 bullets).",
    "Do not add a 'Response' heading.",
    "Use tables only when the user explicitly asks for one or when the content is inherently tabular (labs, ranges, dosing). Otherwise keep the main structure as prose + bullets.",
    "If you include structure, use only these labels: Response, Summary, Plan, Diagnosis, Labs, Treatment, Warnings, Next steps.",
    opts.allowExcerpts
      ? "If quoting slides is requested, include brief verbatim snippets with citations."
      : "Do not include raw excerpt blocks or a 'From the lecture excerpts' section unless explicitly requested.",
  ].join(" ");
}

async function answerWithRetrievedChunks(
  env: Env,
  chunks: Array<{ index: number; text: string }>,
  question: string,
  opts: RetrievedAnswerOptions = {},
): Promise<GeneratedTextResult> {
  const isLibrary = opts.mode === "library";
  const system = isLibrary
    ? buildLibrarySystemPrompt(opts)
    : buildDefaultSystemPrompt(opts);

  const ordered = isLibrary ? [...chunks].sort((a, b) => a.index - b.index) : chunks;
  const cleanedChunks = ordered.map(chunk => ({ ...chunk, text: cleanRetrievedChunkText(chunk.text) }));
  const contextLabel = isLibrary ? "Section" : "Chunk";
  const contextHeader = isLibrary ? "Lecture excerpts" : "Context chunks";
  const contextBody = cleanedChunks.map(chunk => `[${contextLabel} ${chunk.index + 1}]\n${chunk.text}`).join("\n\n");
  const contextBlock = opts.allowExcerpts
    ? `${contextHeader}:\n${contextBody}`
    : `${contextHeader} (REFERENCE — do not quote verbatim unless asked):\n${contextBody}`;
  const continuationNotes = isLibrary
    ? opts.contextNotes || (await buildLibraryContextNotes(env, question, cleanedChunks, { requestId: opts.requestId }))
    : contextBlock;
  const continuationContext = isLibrary
    ? `${contextBlock}\n\n(Internal context notes — do not quote or mimic; keep clinician-style synthesis):\n${continuationNotes || "None"}`
    : contextBlock;

  const result = await runResponsesWithContinuation(env, {
    label: isLibrary ? `library-${opts.requestId || "ask"}` : "ask-retrieved",
    maxOutputTokens: opts.maxOutputTokens ?? ANSWER_MAX_OUTPUT_TOKENS,
    initialText: opts.initialAnswerText || "",
    buildPayload: ({ attempt, accumulatedText }) => {
      const isContinuation = attempt > 1 || Boolean(accumulatedText);
      const contextForAttempt = isContinuation ? continuationContext : contextBlock;
      const userParts: ResponsesInputContent[] = [
        { type: "input_text", text: contextForAttempt },
        { type: "input_text", text: `Question:\n${question}` },
      ];
      if (isContinuation && accumulatedText) {
        userParts.push({ type: "input_text", text: `Partial answer so far:\n${clipAnswerTail(accumulatedText)}` });
      }
      if (isContinuation) {
        userParts.push({
          type: "input_text",
          text: "Continue exactly where you left off. Do not repeat previous sentences or headings. Do not restart tables. Preserve the existing narrative formatting.",
        });
      }

      return {
        model: resolveModelId("gpt-5-mini", env),
        input: [
          { role: "system" as const, content: [{ type: "input_text" as const, text: system }] },
          { role: "user" as const, content: userParts },
        ],
      };
    },
  });

  return {
    ...result,
    text: replaceChunkMarkers(result.text, cleanedChunks),
    fullText: replaceChunkMarkers(result.fullText, cleanedChunks),
  };
}

async function buildLibraryContextNotes(
  env: Env,
  question: string,
  chunks: Array<{ index: number; text: string }>,
  opts: { requestId?: string } = {},
): Promise<string> {
  const compactBody = chunks
    .map(chunk => `[Section ${chunk.index + 1}]\n${cleanRetrievedChunkText(chunk.text)}`)
    .join("\n\n")
    .slice(0, LIBRARY_NOTES_CONTEXT_CHARS);
  if (!compactBody) return "";

  const system = [
    "You are summarizing lecture excerpts into compact context notes for downstream Q&A.",
    "Return concise bullet points or short sentences focused on facts relevant to the question.",
    "Preserve terminology, numeric values, and structural cues (tables/headings) when present.",
    "Do not answer the user's question directly; produce reusable notes only.",
    "These notes are internal scaffolding. Keep them neutral, avoid heavy templating, and do not restate them verbatim in any final answer.",
  ].join(" ");

  const payload = {
    model: resolveModelId("gpt-4o", env),
    max_output_tokens: LIBRARY_NOTES_MAX_OUTPUT_TOKENS,
    input: [
      { role: "system" as const, content: [{ type: "input_text" as const, text: system }] },
      {
        role: "user" as const,
        content: [
          { type: "input_text" as const, text: `Lecture excerpts (trimmed):\n${compactBody}` },
          { type: "input_text" as const, text: `Question: ${question}` },
          {
            type: "input_text" as const,
            text: "Return only the condensed notes (bullet list or short paragraphs). Do not add prefaces.",
          },
        ],
      },
    ],
  };

  try {
    const notes = await callResponsesOnce(env, payload, `library-notes:${opts.requestId || "notes"}`);
    const text = (notes.text || "").trim();
    if (text) return text;
  } catch (err) {
    console.warn("[LIBRARY_NOTES] failed", {
      requestId: opts.requestId,
      error: err instanceof Error ? err.message : String(err),
    });
  }

  return compactBody;
}

type LibraryContinuationState = {
  v: 1;
  docId: string;
  requestId: string;
  question: string;
  contextNotes: string;
  extractedKey?: string;
  options?: {
    allowExcerpts?: boolean;
    tableHeaders?: string[];
    tableOnly?: boolean;
    broad?: boolean;
    listIntent?: boolean;
  };
  answerTail: string;
  segments: number;
  maxOutputTokens: number;
  model: string;
};

function buildLibraryContinuationToken(state: LibraryContinuationState): string {
  const safeState: LibraryContinuationState = {
    ...state,
    contextNotes: (state.contextNotes || "").slice(0, LIBRARY_NOTES_CONTEXT_CHARS),
    answerTail: clipAnswerTail(state.answerTail || ""),
    segments: Math.max(0, state.segments || 0),
    maxOutputTokens: state.maxOutputTokens || ANSWER_MAX_OUTPUT_TOKENS,
    model: state.model || "gpt-5-mini",
    v: 1 as const,
  };
  const payload = JSON.stringify(safeState);
  const bytes = enc.encode(payload);
  let binary = "";
  bytes.forEach(b => {
    binary += String.fromCharCode(b);
  });
  return btoa(binary);
}

function parseLibraryContinuationToken(token: string): LibraryContinuationState | null {
  try {
    const binary = atob(token);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    const decoded = dec.decode(bytes);
    const parsed = JSON.parse(decoded);
    if (
      parsed &&
      parsed.v === 1 &&
      typeof parsed.docId === "string" &&
      typeof parsed.question === "string"
    ) {
      return {
        ...parsed,
        contextNotes: typeof parsed.contextNotes === "string" ? parsed.contextNotes : "",
      } as LibraryContinuationState;
    }
  } catch {
    return null;
  }
  return null;
}

async function handleLibrarySearch(req: Request, env: Env): Promise<Response> {
  const url = new URL(req.url);
  const q = url.searchParams.get("q")?.trim() || "";
  const limitParam = Number(url.searchParams.get("limit"));
  const limit = Number.isFinite(limitParam) && limitParam > 0 ? Math.min(limitParam, 50) : LIBRARY_SEARCH_LIMIT;
  const { bucket } = getLibraryBucket(env);
  const indexRecords = await readLibraryIndex(bucket);
  const ranked = q ? scoreLibraryRecords(q, indexRecords, limit * 2) : indexRecords.slice(0, limit);
  const selected = ranked.slice(0, limit);

  const results: Array<{ docId: string; title: string; bucket: string; key: string; status: "ready" | "missing" | "needs_browser_ocr"; preview?: string; extractedKey: string; }> = [];
  for (const rec of selected) {
    const extractedKey = rec.extractedKey || buildExtractedPath(rec.docId);
    let status: "ready" | "missing" | "needs_browser_ocr" = rec.status || "missing";
    try {
      const head = typeof bucket.head === "function"
        ? await bucket.head(extractedKey)
        : await bucket.get(extractedKey, { range: { offset: 0, length: 0 } as any });
      if (head) status = "ready";
    } catch {
      status = status === "needs_browser_ocr" ? "needs_browser_ocr" : "missing";
    }
    results.push({
      docId: rec.docId,
      title: rec.title,
      bucket: rec.bucket,
      key: rec.key,
      status,
      preview: normalizePreview(rec.preview),
      extractedKey,
    });
  }

  return json({ results });
}

async function handleLibraryList(req: Request, env: Env): Promise<Response> {
  const url = new URL(req.url);
  const prefix = url.searchParams.get("prefix") || "library/";
  const { bucket } = getLibraryBucket(env);
  const indexRecords = await readLibraryIndex(bucket);
  const filtered = indexRecords
    .filter(rec => rec.key?.startsWith(prefix) && isPdfKey(rec.key || ""))
    .sort((a, b) => (a.title || "").localeCompare(b.title || ""));
  const items = await Promise.all(
    filtered.map(async rec => {
      const extractedKey = rec.extractedKey || buildExtractedPath(rec.docId);
      let status: "ready" | "missing" | "needs_browser_ocr" = rec.status || "missing";
      try {
        const head = typeof bucket.head === "function"
          ? await bucket.head(extractedKey)
          : await bucket.get(extractedKey, { range: { offset: 0, length: 0 } as any });
        if (head) status = "ready";
      } catch {
        status = status === "needs_browser_ocr" ? "needs_browser_ocr" : "missing";
      }
      return {
        title: rec.title,
        bucket: rec.bucket,
        key: rec.key,
        docId: rec.docId,
        extractedKey,
        status,
      };
    }),
  );
  return json({ items });
}

async function handleLibraryIngest(req: Request, env: Env): Promise<Response> {
  const body = await req.json().catch(() => null);
  if (!body || typeof body !== "object") {
    return json({ error: "Send JSON { bucket, key }." }, 400);
  }
  const bucketName = typeof (body as any).bucket === "string" ? (body as any).bucket.trim() : "";
  const key = typeof (body as any).key === "string" ? (body as any).key.trim() : "";
  if (!bucketName || !key) return json({ error: "Missing bucket or key." }, 400);

  try {
    const result = await ingestLibraryObject(env, { bucketName, key });
    return json(result);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    const status = (err as any)?.statusCode && Number.isInteger((err as any).statusCode) ? Number((err as any).statusCode) : 500;
    console.error("[LIBRARY-INGEST] failed", { bucketName, key, error: msg });
    return json({ status: "error", stage: (err as any)?.stage || "unknown", message: msg }, status);
  }
}

async function handleLibraryAsk(req: Request, env: Env): Promise<Response> {
  const requestId = `lib-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const debugLectureAsk = (env as any)?.DEBUG_LECTURE_ASK === "1" || req.headers.get("x-debug-lecture-ask") === "1";
  const body = await req.json().catch(() => null);
  if (!body || typeof body !== "object") {
    return jsonNoStore({ error: "Send JSON { docId, question }.", requestId }, 400);
  }
  const docId = typeof (body as any).docId === "string" ? (body as any).docId.trim() : "";
  const question = typeof (body as any).question === "string" ? (body as any).question.trim() : "";
  const clientReqId = typeof (body as any).requestId === "string" ? (body as any).requestId.trim() : "";
  const effectiveReqId = clientReqId || requestId;
  if (!docId || !question) return jsonNoStore({ error: "Missing docId or question.", requestId: effectiveReqId }, 400);

  const { bucket } = getLibraryBucket(env);
  let docTitle: string | undefined;
  try {
    const indexRecords = await readLibraryIndex(bucket);
    const match = indexRecords.find(rec => rec.docId === docId);
    docTitle = match?.title || undefined;
  } catch {
    docTitle = undefined;
  }
  try {
    const manifest = await readManifest(bucket, docId);
    if (manifest?.title && !docTitle) {
      docTitle = manifest.title;
    }
  } catch {
    // best-effort manifest lookup
  }
  const cache = await loadCachedExtraction(bucket, docId);
  const extractedText =
    cache.text ||
    (cache as any).extractedText ||
    (cache as any).content ||
    (cache as any).ocrText ||
    (cache as any).textContent ||
    null;
  if (!extractedText) {
    console.error("[LIBRARY-ASK] missing extraction", { requestId: effectiveReqId, docId });
    return jsonNoStore({
      v: 1,
      ok: false,
      error: "extracted_text_not_found",
      details: "No cached extraction for this docId.",
      requestId: effectiveReqId,
      docId,
      answer: "Lecture text not found for this selection.",
    }, 404);
  }
  const text = extractedText.slice(0, MAX_OCR_TEXT_LENGTH);
  if (debugLectureAsk) {
    console.log("[LIBRARY_ASK_START]", { docId, requestId: effectiveReqId, questionLen: question.length });
  }
  console.log("[LIBRARY_CONTEXT]", { requestId: effectiveReqId, docId, extractedKey: cache.extractedKey, contextLen: text.length });
  console.log("[LIB_ASK_GUARD]", { req: effectiveReqId, docId, hasKey: Boolean(cache.extractedKey), r2TextLen: text.length });
  if (text.length < 300) {
    console.error("[LIBRARY_CONTEXT_MISSING]", { docId, contextLen: text.length, requestId: effectiveReqId });
    return jsonNoStore({
      v: 1,
      ok: false,
      error: { code: "ocr_not_ready", message: "Extracted text too small or not ready." },
      answer: "OCR is still processing or extracted text is not ready yet. Please retry in a moment or re-ingest.",
      docId,
      requestId: effectiveReqId,
    });
  }
  const chunks = chunkText(text, { size: RETRIEVAL_CHUNK_SIZE, overlap: RETRIEVAL_CHUNK_OVERLAP });
  const intent = detectLibraryIntent(question);
  const excerptsAllowed = intent.wantsTable ? false : allowExcerpts(question);
  let selected = selectLibraryChunks(question, chunks, intent);
  if (!selected.length) {
    selected = chunks.slice(0, Math.min(RETRIEVAL_TOP_K, chunks.length));
  }
  const selectedCharCount = selected.reduce((sum, chunk) => sum + chunk.text.length, 0);
  console.log("[LIBRARY-ASK] retrieval", {
    requestId: effectiveReqId,
    docId,
    totalChunks: chunks.length,
    selected: selected.map(c => c.index),
    mode: intent.isBroad ? "broad" : "narrow",
    chars: selectedCharCount,
    extractedLen: text.length,
    extractedKey: cache.extractedKey,
  });

  try {
    debugFormatLog("[LIBRARY] format route", {
      docId,
      wantsTable: intent.wantsTable,
      broad: intent.isBroad,
      allowExcerpts: excerptsAllowed,
      selectedChunks: selected.length,
    });
    const contextNotes = await buildLibraryContextNotes(env, question, selected, { requestId: effectiveReqId });
    const rawAnswerResult = await answerWithRetrievedChunks(env, selected, question, {
      mode: "library",
      broad: intent.isBroad,
      listIntent: intent.listIntent,
      tableSchema: intent.wantsTable ? undefined : undefined,
      tableOnly: intent.wantsTable,
      tableHeaders: intent.wantsTable ? deriveTableHeaders(question) : undefined,
      allowExcerpts: excerptsAllowed,
      requestId: effectiveReqId,
      contextNotes,
    });
    const rawAnswer = rawAnswerResult.fullText;
    if (!rawAnswer.trim()) {
      console.error("[LIBRARY_ASK] empty before truncation check", { requestId: effectiveReqId, docId });
      return jsonNoStore({
        v: 1,
        ok: false,
        error: "empty_answer",
        requestId: effectiveReqId,
        docId,
        answer: "Response generation timed out before any content was produced. Please retry.",
      }, 200);
    }
    const inherentlyTabular = isInherentlyTabularQuestion(question) || intent.wantsConditionLabsTable;
    const preferTable = intent.wantsTable || inherentlyTabular;
    const formatted = preferTable ? trimAfterFirstTable(rawAnswer) : rawAnswer;
    const repaired = preferTable ? ensureTablePresence(formatted, deriveTableHeaders(question)) : formatted;
    const sanitized = sanitizeAnswer(repaired, { allowExcerpts: excerptsAllowed });
    const finalCore = finalizeAssistantText(rawAnswer, sanitized);
    const finalAnswer = rawAnswerResult.truncated ? appendTruncationNotice(finalCore) : finalCore;
    const answerLength = (finalAnswer || "").trim().length;
    const evidence = excerptsAllowed ? buildEvidenceFromChunks(selected, docId) : undefined;
    console.log("[LIBRARY_ASK] answerLength", answerLength, "evidenceCount", evidence?.length || 0, "req", effectiveReqId, "docId", docId);
    if (debugLectureAsk) {
      console.log("[LIBRARY_ASK][DEBUG]", {
        requestId: effectiveReqId,
        promptLength: question.length,
        maxOutputTokens: ANSWER_MAX_OUTPUT_TOKENS,
        listIntent: intent.listIntent,
        chunksCount: chunks.length,
        selected: selected.map(c => c.index),
      });
    }
    if (intent.listIntent) {
      console.log("[LIST_INTENT]", { requestId: effectiveReqId, docId, matched: true, templateApplied: false });
    }
    if (!answerLength) {
      console.error("[LIBRARY_ASK] empty answer", { requestId: effectiveReqId, docId, selected: selected.map(c => c.index) });
      return jsonNoStore({
        v: 1,
        ok: false,
        error: "empty_answer",
        requestId: effectiveReqId,
        docId,
        answer: `No answer text was produced for this prompt. Try rephrasing or ask for fewer items. (req: ${effectiveReqId})`,
      }, 200);
    }
    debugFormatLog("[LIBRARY] format result", {
      preferTable,
      rawLength: rawAnswer.length,
      repairedLength: repaired.length,
      finalLength: finalAnswer.length,
    });
    const references = buildReferencesFromChunks(selected, docId);
    const payload: Record<string, unknown> = {
      v: 1,
      ok: true,
      answer: finalAnswer,
      docId,
      extractedKey: cache.extractedKey,
      references,
      meta: docTitle ? { docTitle } : undefined,
      requestId: effectiveReqId,
    };
    if (rawAnswerResult.truncated && answerLength > 0) {
      payload.truncated = true;
      payload.continuationToken = buildLibraryContinuationToken({
        docId,
        requestId: effectiveReqId,
        question,
        contextNotes,
        extractedKey: cache.extractedKey,
        options: {
          allowExcerpts: excerptsAllowed,
          tableHeaders: intent.wantsTable ? deriveTableHeaders(question) : undefined,
          tableOnly: intent.wantsTable,
          broad: intent.isBroad,
          listIntent: intent.listIntent,
        },
        answerTail: clipAnswerTail(finalCore),
        segments: rawAnswerResult.attempts,
        maxOutputTokens: ANSWER_MAX_OUTPUT_TOKENS,
        model: resolveModelId("gpt-5-mini", env),
      });
    }
    if (evidence && evidence.length) {
      payload.evidence = evidence;
    }
    const analyticsTags = filterMetadata([question], MAX_TAGS);
    const analyticsEntities = filterMetadata([question], MAX_TAGS);
    const lectureAnchors = references.map(ref => ref.title).filter(Boolean) as string[];
    try {
      const event = buildAnalyticsEvent({
        lectureId: docId,
        lectureTitle: docTitle || null,
        question,
        docId,
        tags: analyticsTags,
        entities: analyticsEntities,
        lectureAnchors,
        model: resolveModelId("gpt-5-mini", env),
      });
      await writeAnalyticsEvent(env, event);
    } catch (err) {
      console.error("[ANALYTICS_WRITE_FAILED]", {
        docId,
        error: err instanceof Error ? err.message : String(err),
      });
    }
    if (debugLectureAsk) {
      console.log("[LIB_ASK_OUT]", { req: effectiveReqId, ok: true, answerLen: answerLength, refs: references.length });
    }
    console.log("[ASK_END]", {
      req: effectiveReqId,
      ok: true,
      answerLen: answerLength,
      truncated: Boolean(payload.truncated),
      hasToken: Boolean(payload.continuationToken),
    });
    return jsonNoStore(payload);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[LIBRARY-ASK] failed", { docId, error: msg, requestId: effectiveReqId });
    return jsonNoStore({
      v: 1,
      ok: false,
      error: "answer_failed",
      details: msg,
      requestId: effectiveReqId,
      docId,
      answer: `Error generating answer: ${msg}`,
    }, 200);
  }
}

async function handleLibraryAskContinue(req: Request, env: Env): Promise<Response> {
  const requestId = `lib-cont-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const body = await req.json().catch(() => null);
  if (!body || typeof body !== "object") {
    return jsonNoStore({ ok: false, error: "invalid_body", message: "Send JSON { docId, continuationToken }.", requestId }, 400);
  }
  const docId = typeof (body as any).docId === "string" ? (body as any).docId.trim() : "";
  const tokenInput = typeof (body as any).continuationToken === "string" ? (body as any).continuationToken.trim() : "";
  const clientReqId = typeof (body as any).requestId === "string" ? (body as any).requestId.trim() : "";
  const effectiveReqId = clientReqId || requestId;
  if (!docId || !tokenInput) {
    return jsonNoStore({
      ok: false,
      error: "missing_params",
      message: "docId and continuationToken are required.",
      requestId: effectiveReqId,
    }, 400);
  }

  const parsed = parseLibraryContinuationToken(tokenInput);
  if (!parsed || parsed.docId !== docId) {
    return jsonNoStore({
      ok: false,
      error: "invalid_continuation_token",
      message: "Continuation token is invalid or does not match the requested docId.",
      requestId: effectiveReqId,
    }, 400);
  }

  if ((parsed.segments || 0) >= LIBRARY_TOTAL_CONTINUATIONS) {
    return jsonNoStore({
      v: 1,
      ok: false,
      error: "continuation_limit",
      message: "Continuation limit reached before more content was produced. Please retry the question.",
      docId,
      requestId: parsed.requestId || effectiveReqId,
    }, 200);
  }

  const { bucket } = getLibraryBucket(env);
  const cache = await loadCachedExtraction(bucket, docId);
  const extractedText =
    cache.text ||
    (cache as any).extractedText ||
    (cache as any).content ||
    (cache as any).ocrText ||
    (cache as any).textContent ||
    null;
  if (!extractedText) {
    return jsonNoStore({
      v: 1,
      ok: false,
      error: "extracted_text_not_found",
      details: "No cached extraction for this docId.",
      requestId: effectiveReqId,
      docId,
      answerSegment: "Lecture text not found for this selection.",
    }, 404);
  }
  const text = extractedText.slice(0, MAX_OCR_TEXT_LENGTH);
  if (text.length < 300) {
    return jsonNoStore({
      v: 1,
      ok: false,
      error: { code: "ocr_not_ready", message: "Extracted text too small or not ready." },
      answerSegment: "OCR is still processing or extracted text is not ready yet. Please retry in a moment or re-ingest.",
      docId,
      requestId: effectiveReqId,
    });
  }
  if (parsed.extractedKey && parsed.extractedKey !== cache.extractedKey) {
    return jsonNoStore({
      ok: false,
      error: "stale_continuation_token",
      message: "The lecture was re-processed; please restart the ask.",
      docId,
      requestId: effectiveReqId,
    }, 409);
  }

  const options = parsed.options || {};
  let contextNotes = parsed.contextNotes || "";
  const chunks = chunkText(text, { size: RETRIEVAL_CHUNK_SIZE, overlap: RETRIEVAL_CHUNK_OVERLAP });
  const intent = detectLibraryIntent(parsed.question);
  const selected = selectLibraryChunks(parsed.question, chunks, intent);
  if (!contextNotes) {
    contextNotes = await buildLibraryContextNotes(env, parsed.question, selected, { requestId: parsed.requestId || effectiveReqId });
  }
  if (!contextNotes) {
    contextNotes = "Context notes unavailable; continue the previous answer based on the question and prior text.";
  }

  const contextLabel = "Section";
  const contextHeader = "Lecture excerpts";
  const contextBody = selected
    .map(chunk => `[${contextLabel} ${chunk.index + 1}]\n${cleanRetrievedChunkText(chunk.text)}`)
    .join("\n\n");
  const contextBlock = options.allowExcerpts
    ? `${contextHeader}:\n${contextBody}`
    : `${contextHeader} (REFERENCE — do not quote verbatim unless asked):\n${contextBody}`;
  const continuationContext = `${contextBlock}\n\n(Internal context notes — do not quote or mimic; keep clinician-style synthesis):\n${contextNotes || "None"}`;

  const system = buildLibrarySystemPrompt({
    mode: "library",
    allowExcerpts: options.allowExcerpts,
    tableHeaders: options.tableHeaders,
    tableOnly: options.tableOnly,
    broad: options.broad,
    listIntent: options.listIntent,
  });

  const remainingBudget = Math.max(1, LIBRARY_TOTAL_CONTINUATIONS - (parsed.segments || 0));
  const maxAttempts = Math.min(ANSWER_MAX_CONTINUATIONS + 1, remainingBudget);
  const model = parsed.model || resolveModelId("gpt-5-mini", env);

  const result = await runResponsesWithContinuation(env, {
    label: `library-continue-${parsed.requestId || effectiveReqId}`,
    initialText: parsed.answerTail || "",
    maxAttempts,
    maxOutputTokens: parsed.maxOutputTokens || ANSWER_MAX_OUTPUT_TOKENS,
    buildPayload: ({ accumulatedText }) => {
      const userParts: ResponsesInputContent[] = [
        { type: "input_text", text: continuationContext },
        { type: "input_text", text: `Question:\n${parsed.question}` },
      ];
      if (accumulatedText) {
        userParts.push({
          type: "input_text",
          text: `Partial answer so far:\n${clipAnswerTail(accumulatedText)}`,
        });
      }
      userParts.push({
        type: "input_text",
        text: "Continue exactly where you left off. Do not repeat earlier sentences or headings. Do not restart tables. Preserve the existing narrative formatting.",
      });
      return {
        model,
        input: [
          { role: "system" as const, content: [{ type: "input_text" as const, text: system }] },
          { role: "user" as const, content: userParts },
        ],
      };
    },
  });

  const updatedSegments = (parsed.segments || 0) + result.attempts;
  const continuationAllowed = updatedSegments < LIBRARY_TOTAL_CONTINUATIONS;
  let continuationToken: string | undefined;
  const segmentText = (result.text || "").trim();
  if (!segmentText) {
    return jsonNoStore({
      v: 1,
      ok: false,
      error: "empty_answer",
      message: "Response generation timed out before any content was produced. Please retry.",
      requestId: parsed.requestId || effectiveReqId,
      docId,
    }, 200);
  }
  if (result.truncated && continuationAllowed) {
    continuationToken = buildLibraryContinuationToken({
      ...parsed,
      requestId: parsed.requestId || effectiveReqId,
      contextNotes,
      answerTail: clipAnswerTail(result.fullText),
      segments: updatedSegments,
      maxOutputTokens: parsed.maxOutputTokens || ANSWER_MAX_OUTPUT_TOKENS,
      model,
    });
  }

  const answerSegment = result.truncated && segmentText ? appendTruncationNotice(segmentText) : segmentText;

  const responsePayload = {
    v: 1,
    ok: true,
    answerSegment,
    docId,
    requestId: parsed.requestId || effectiveReqId,
    continuationToken,
    done: !continuationToken,
    truncated: result.truncated,
    limitReached: result.truncated && !continuationToken,
  };
  console.log("[ASK_END]", {
    req: parsed.requestId || effectiveReqId,
    ok: true,
    answerLen: answerSegment.length,
    truncated: result.truncated,
    hasToken: Boolean(continuationToken),
  });
  return jsonNoStore(responsePayload);
}

async function handleLibraryBatchIndex(req: Request, env: Env): Promise<Response> {
  const body = await req.json().catch(() => ({}));
  const bucketsInput = Array.isArray((body as any).buckets) ? (body as any).buckets : null;
  const bucketNames = bucketsInput?.length
    ? bucketsInput.map((b: unknown) => String(b)).filter(Boolean)
    : Object.keys(BUCKET_BINDINGS);

  const { bucket: libraryBucket } = getLibraryBucket(env);
  const records: LibraryIndexRecord[] = [];
  let totalObjects = 0;
  let indexed = 0;

  for (const bucketName of bucketNames) {
    let lookup: R2Bucket;
    try {
      lookup = getBucketByName(env, bucketName);
    } catch {
      console.warn("[LIBRARY-BATCH-INDEX] unknown bucket", { bucketName });
      continue;
    }
    const objects = await listPdfObjectsFromBucket(lookup);
    totalObjects += objects.length;
    for (const obj of objects) {
      const canonicalKey = typeof obj.key === "string" ? obj.key.trim() : obj.key;
      const { docId, basis, fieldsUsed, uploaded } = await computeDocId(bucketName, canonicalKey, {
        etag: obj.etag,
        size: obj.size,
        uploaded: obj.uploaded,
      });
      const title = titleFromKey(canonicalKey);
      const extractedKey = buildExtractedPath(docId);
      let status: "ready" | "missing" | "needs_browser_ocr" = "missing";
      const manifest = await readManifest(libraryBucket, docId);
      if (manifest && manifest.method === "ocr" && !manifest.pagesProcessed) {
        status = "needs_browser_ocr";
      }
      try {
        const head = typeof libraryBucket.head === "function" ? await libraryBucket.head(extractedKey) : await libraryBucket.get(extractedKey, { range: { offset: 0, length: 0 } as any });
        if (head) status = "ready";
      } catch {
        status = "missing";
      }
      const record: LibraryIndexRecord = {
        docId,
        bucket: bucketName,
        key: canonicalKey,
        title,
        normalizedTokens: tokensFromTitle(title),
        hashBasis: basis,
        hashFieldsUsed: fieldsUsed,
        etag: obj.etag,
        size: obj.size,
        uploaded: uploaded || obj.uploaded?.toISOString?.(),
        status,
        extractedKey,
        manifestKey: buildManifestPath(docId),
        preview: manifest?.preview,
      };
      records.push(record);
      indexed += 1;
    }
  }

  if (records.length) {
    await writeLibraryIndex(libraryBucket, records);
    await Promise.all(
      records.map(rec =>
        libraryBucket.put(buildIndexKeyForDoc(rec.docId), JSON.stringify(rec), {
          httpMetadata: { contentType: "application/json; charset=utf-8" },
        }),
      ),
    );
  }

  return json({
    status: "ok",
    buckets: bucketNames,
    indexed,
    totalObjects,
    indexKey: LIBRARY_INDEX_KEY,
  });
}

async function handleLibraryBatchIngest(req: Request, env: Env): Promise<Response> {
  const body = await req.json().catch(() => ({}));
  const bucketsInput = Array.isArray((body as any).buckets) ? (body as any).buckets : null;
  const modeInput = typeof (body as any).mode === "string" ? (body as any).mode : "embedded_only";
  const mode = modeInput === "full" ? "full" : "embedded_only";
  const maxDocsRaw = Number((body as any).maxDocs);
  const maxDocs = Number.isFinite(maxDocsRaw) && maxDocsRaw > 0 ? maxDocsRaw : undefined;
  const bucketNames = bucketsInput?.length
    ? bucketsInput.map((b: unknown) => String(b)).filter(Boolean)
    : Object.keys(BUCKET_BINDINGS);

  const queue: Array<{ docId: string; bucket: string; key: string; title: string; pageCount?: number; downloadUrl: string }> = [];
  let processed = 0;
  let ready = 0;
  let skipped = 0;
  let needsBrowser = 0;

  for (const bucketName of bucketNames) {
    let lookup: R2Bucket;
    try {
      lookup = getBucketByName(env, bucketName);
    } catch {
      console.warn("[LIBRARY-BATCH-INGEST] unknown bucket", { bucketName });
      continue;
    }
    const objects = await listPdfObjectsFromBucket(lookup, { limit: maxDocs ? Math.max(1, maxDocs - processed) : undefined });
    for (const obj of objects) {
      if (maxDocs && processed >= maxDocs) break;
      processed += 1;
      try {
        const result = await ingestLibraryObject(env, { bucketName, key: obj.key, skipCache: false, mode });
        if (result.status === "cache_hit") {
          skipped += 1;
          continue;
        }
        if (result.status === "needs_browser_ocr") {
          needsBrowser += 1;
          if (mode === "full" && queue.length < LIBRARY_QUEUE_MAX) {
            queue.push({
              docId: result.docId,
              bucket: bucketName,
              key: obj.key,
              title: titleFromKey(obj.key),
              pageCount: result.pageCount,
              downloadUrl: result.downloadUrl || buildLibraryDownloadUrl(bucketName, obj.key),
            });
          }
          continue;
        }
        ready += 1;
      } catch (err) {
        console.warn("[LIBRARY-BATCH-INGEST] failed for object", { bucketName, key: obj.key, error: err instanceof Error ? err.message : String(err) });
      }
    }
  }

  let queueKey = "";
  if (queue.length) {
    const { bucket } = getLibraryBucket(env);
    queueKey = `${LIBRARY_QUEUE_PREFIX}${Date.now()}.json`;
    await bucket.put(queueKey, JSON.stringify({ docs: queue, createdAt: new Date().toISOString() }), {
      httpMetadata: { contentType: "application/json; charset=utf-8" },
    });
  }

  return json({
    status: "ok",
    mode,
    processed,
    ready,
    skipped,
    needs_browser_ocr: needsBrowser,
    queueKey: queueKey || undefined,
    queued: queue.length,
  });
}

async function handleSignedUrl(req: Request, env: Env): Promise<Response> {
  const url = new URL(req.url);
  const bucketName = url.searchParams.get("bucket") || "";
  const key = url.searchParams.get("key") || "";
  if (!bucketName || !key) {
    return json({ error: "Missing bucket or key." }, 400);
  }
  try {
    const bucket = getBucketByName(env, bucketName);
    const exists = typeof bucket.head === "function" ? await bucket.head(key) : await bucket.get(key, { range: { offset: 0, length: 0 } as any });
    if (!exists) return json({ error: "Not found." }, 404);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return json({ error: "Not found", details: msg }, 404);
  }
  const downloadUrl = buildLibraryDownloadUrl(bucketName, key);
  return json({ url: downloadUrl, bucket: bucketName, key });
}

async function handleAskDoc(req: Request, env: Env): Promise<Response> {
  const body = await req.json().catch(() => null);
  if (!body || typeof body !== "object") {
    return jsonNoStore({ error: "Send JSON { extractedKey?, fileHash?, question }." }, 400);
  }
  const question = typeof (body as any).question === "string" ? (body as any).question.trim() : "";
  const extractedKeyInput = typeof (body as any).extractedKey === "string" ? (body as any).extractedKey.trim() : "";
  const fileHash = typeof (body as any).fileHash === "string" ? (body as any).fileHash.trim() : "";
  if (!question) return jsonNoStore({ error: "Missing question." }, 400);
  if (!extractedKeyInput && !fileHash) return jsonNoStore({ error: "Provide extractedKey or fileHash." }, 400);

  const extractedKey = extractedKeyInput || buildExtractedKeyForHash(fileHash);
  const { bucket } = getExtractionBucket(env);
  const object = await bucket.get(extractedKey);
  if (!object || !object.body) {
    return jsonNoStore({ error: "extracted_text_not_found", details: "No extracted text stored for this file." }, 404);
  }
  const text = normalizePlainText(await object.text()).slice(0, MAX_OCR_TEXT_LENGTH);
  if (!text) {
    return jsonNoStore({ error: "extracted_text_empty", details: "Stored text is empty." }, 404);
  }

  const chunks = chunkText(text, { size: RETRIEVAL_CHUNK_SIZE, overlap: RETRIEVAL_CHUNK_OVERLAP });
  const ranked = rankChunks(question, chunks, RETRIEVAL_TOP_K);
  const selected = ranked.length ? ranked : chunks.slice(0, Math.min(RETRIEVAL_TOP_K, chunks.length));
  console.log("[ASK-DOC] retrieval", {
    extractedKey,
    totalChunks: chunks.length,
    selectedChunks: selected.map(chunk => chunk.index),
  });

  try {
    const tableIntent = isTableIntent(question);
    const excerptsAllowed = tableIntent ? false : allowExcerpts(question);
    const tableHeaders = tableIntent ? deriveTableHeaders(question) : undefined;
    debugFormatLog("[ASK-DOC] format route", {
      tableIntent,
      allowExcerpts: excerptsAllowed,
      tableHeaders,
      selectedChunks: selected.length,
    });
    const rawAnswerResult = await answerWithRetrievedChunks(env, selected, question, {
      tableOnly: tableIntent,
      tableSchema: tableIntent ? undefined : undefined,
      tableHeaders,
      allowExcerpts: excerptsAllowed,
    });
    const rawAnswer = rawAnswerResult.fullText;
    const preferTable = tableIntent || isInherentlyTabularQuestion(question);
    const formatted = preferTable ? trimAfterFirstTable(rawAnswer) : rawAnswer;
    const repaired = preferTable ? ensureTablePresence(formatted, tableHeaders || deriveTableHeaders(question)) : formatted;
    const sanitized = sanitizeAnswer(repaired, { allowExcerpts: excerptsAllowed });
    const answerCore = finalizeAssistantText(rawAnswer, sanitized);
    const answer = rawAnswerResult.truncated ? appendTruncationNotice(answerCore) : answerCore;
    debugFormatLog("[ASK-DOC] format result", {
      preferTable,
      rawLength: rawAnswer.length,
      repairedLength: repaired.length,
      finalLength: answer.length,
    });
    const references = buildReferencesFromChunks(selected);
    const evidence = excerptsAllowed ? buildEvidenceFromChunks(selected) : undefined;
    const payload: Record<string, unknown> = {
      answer,
      method: "retrieval",
      extractedKey,
      chunksUsed: selected.map(chunk => chunk.index),
      references,
    };
    if (evidence && evidence.length) {
      payload.evidence = evidence;
    }
    return jsonNoStore({
      ...payload,
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[ASK-DOC] Answer failed", err);
    return jsonNoStore({ error: "answer_failed", details: msg }, 502);
  }
}

async function handleOcrPage(req: Request, env: Env): Promise<Response> {
  const contentType = req.headers.get("content-type") || "";
  if (!contentType.includes("multipart/form-data")) {
    return json({ error: "Send multipart/form-data with fields: fileId, pageIndex, image" }, 400);
  }
  const form = await req.formData();
  const fileId = typeof form.get("fileId") === "string" ? (form.get("fileId") as string).trim() : "";
  const fileHash = typeof form.get("fileHash") === "string" ? (form.get("fileHash") as string).trim() : "";
  const rawIndex = form.get("pageIndex");
  const pageIndex = typeof rawIndex === "string" ? Number(rawIndex) : Number(rawIndex);
  const image = form.get("image");

  if (!fileId && !fileHash) return json({ error: "Missing fileId or fileHash." }, 400);
  if (!Number.isFinite(pageIndex)) return json({ error: "Invalid pageIndex." }, 400);
  if (!(image instanceof File)) return json({ error: "Missing image file." }, 400);

  try {
    const mimeType = image.type || "image/png";
    const bytes = new Uint8Array(await image.arrayBuffer());
    const prepared = prepareOcrImageInput(bytes, mimeType);
    if (!prepared.ok) {
      console.warn("[OCR-PAGE] Invalid image bytes", {
        fileHash: fileHash || fileId,
        pageIndex,
        head: prepared.head,
        mimeType: prepared.mimeType,
      });
      return json(invalidImagePayload({ pageIndex }), 422);
    }
    logOcrInputDebug({
      key: fileHash || fileId || "ocr-page",
      head: prepared.head,
      mimeType: prepared.mimeType,
      signature: prepared.signature,
      byteLength: prepared.byteLength,
    });
    const label = Number.isFinite(pageIndex) ? `Page ${Number(pageIndex) + 1}` : "Page";
    const started = Date.now();
    const prompt = "You are an OCR service. Transcribe every readable character from the page image. Return plain text only.";
    const text = await callResponsesOcr(
      env,
      [
        { type: "input_text", text: prompt },
        { type: "input_image", image_url: prepared.dataUrl, detail: "high" as InputImageDetail },
      ],
      { label: `ocr-page:${fileHash || fileId}:${pageIndex}`, maxOutputTokens: OCR_PAGE_OUTPUT_TOKENS },
    );
    const normalized = normalizeExtractedText(text).slice(0, MAX_OCR_TEXT_LENGTH);
    if (!normalized) {
      return json({ error: "ocr_empty", details: "No text extracted for this page." }, 422);
    }
    console.log("[OCR-PAGE] completed", {
      fileHash: fileHash || fileId,
      pageIndex,
      bytes: bytes.length,
      ms: Date.now() - started,
      length: normalized.length,
    });
    return json({ pageIndex: Number(pageIndex), text: normalized });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[OCR-PAGE] Failed", { fileHash: fileHash || fileId, pageIndex, error: msg });
    return json({ error: "ocr_failed", details: msg }, 502);
  }
}

async function handleOcrFinalize(req: Request, env: Env): Promise<Response> {
  const body = await req.json().catch(() => null);
  if (!body || typeof body !== "object") {
    return json({ error: "Send JSON { fileId?, fileHash?, pages: [{ pageIndex, text }] }." }, 400);
  }
  const fileId = typeof (body as any).fileId === "string" ? (body as any).fileId.trim() : "";
  const fileHash = typeof (body as any).fileHash === "string" ? (body as any).fileHash.trim() : "";
  const filename = typeof (body as any).filename === "string" ? (body as any).filename : undefined;
  const pages = Array.isArray((body as any).pages) ? (body as any).pages : [];
  const totalPages = Number((body as any).totalPages);
  if (!fileId && !fileHash) return json({ error: "Missing fileId or fileHash." }, 400);
  if (!pages.length) return json({ error: "No pages provided." }, 400);

  const normalizedPages = pages
    .map((entry: any) => {
      const idxRaw = typeof entry?.pageIndex === "number" ? entry.pageIndex : Number(entry?.pageIndex);
      const text = typeof entry?.text === "string" ? entry.text : "";
      if (!Number.isFinite(idxRaw)) return null;
      const clean = normalizeExtractedText(text);
      if (!clean) return null;
      return { pageIndex: Number(idxRaw), text: clean };
    })
    .filter(
      (p: { pageIndex: number; text: string } | null): p is { pageIndex: number; text: string } => Boolean(p),
    )
    .sort((a: { pageIndex: number; text: string }, b: { pageIndex: number; text: string }) => a.pageIndex - b.pageIndex);

  if (!normalizedPages.length) {
    return json({ error: "No text to finalize.", details: "All pages were empty." }, 422);
  }

  const normalizedResult = normalizePages(normalizedPages);
  const combined = normalizedResult.text.slice(0, MAX_OCR_TEXT_LENGTH);

  if (!combined.trim()) {
    return json({ error: "finalize_empty", details: "Combined OCR text is empty." }, 422);
  }

  const { bucket } = getExtractionBucket(env);
  const extractedKey = fileHash ? buildExtractedKeyForHash(fileHash) : buildExtractedKey(fileId);
  const existingPages = new Map<number, string>();
  let manifest: PdfManifest | null = null;
  if (fileHash) {
    manifest = await readManifest(bucket, fileHash);
  }
  try {
    const existing = await bucket.get(extractedKey);
    if (existing && existing.body) {
      const existingText = normalizePlainText(await existing.text());
      parseNormalizedPagesFromText(existingText).forEach((text, pageIdx) => existingPages.set(pageIdx, text));
    }
  } catch {}

  normalizedPages.forEach(page => {
    existingPages.set(page.pageIndex, page.text.trim());
  });

  const rebuilt = Array.from(existingPages.entries())
    .sort((a, b) => a[0] - b[0])
    .map(([pageIndex, text]) => `--- Page ${pageIndex + 1} ---\n${text}`)
    .join("\n\n");

  const normalizedFinal = normalizePlainText(rebuilt).slice(0, MAX_OCR_TEXT_LENGTH);

  try {
    await bucket.put(extractedKey, normalizedFinal, {
      httpMetadata: { contentType: "text/plain; charset=utf-8" },
    });
    const preview = normalizedFinal.slice(0, ONE_SHOT_PREVIEW_LIMIT);
    if (fileHash) {
      const newRanges = mergeRanges([
        ...(manifest?.ranges || []),
        ...buildRangesFromPages(normalizedPages.map(p => p.pageIndex)),
      ]);
      const pagesProcessed = newRanges.reduce((total, range) => total + (range.end - range.start + 1), 0);
      const nextManifest: PdfManifest = {
        fileHash,
        filename: manifest?.filename || filename,
        method: "ocr",
        ocrStatus: "finalized",
        pagesProcessed,
        pageCount: Number.isFinite(totalPages) ? Number(totalPages) : manifest?.pageCount,
        ranges: newRanges,
        createdAt: manifest?.createdAt || new Date().toISOString(),
        preview,
        extractedKey,
        updatedAt: new Date().toISOString(),
      };
      await writeManifest(bucket, nextManifest);
      console.log("[OCR-FINALIZE] persisted OCR text", {
        fileHash,
        pagesProcessed,
        pageCount: nextManifest.pageCount,
        ranges: newRanges,
      });
    }
    return json({
      extractionStatus: "ok",
      method: "ocr",
      extractedKey,
      preview,
      fileId,
      fileHash,
      pageCount: Number.isFinite(totalPages) ? Number(totalPages) : undefined,
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[OCR-FINALIZE] Failed to persist OCR text", { fileId: fileHash || fileId, error: msg });
    return json({ error: "persist_failed", details: msg }, 500);
  }
}

async function ingestLibraryObject(
  env: Env,
  {
    bucketName,
    key,
    skipCache = false,
    mode = "single",
  }: { bucketName: string; key: string; skipCache?: boolean; mode?: "embedded_only" | "full" | "single" },
) {
  const stageError = (stage: string, message: string, statusCode = 500) => {
    const err = new Error(message) as Error & { stage?: string; statusCode?: number };
    err.stage = stage;
    err.statusCode = statusCode;
    return err;
  };

  const canonicalKey = key.trim();
  const sourceLookup = lookupBucket(env, bucketName) || lookupBucket(env, BUCKET_BINDINGS[bucketName as keyof typeof BUCKET_BINDINGS]);
  if (!sourceLookup) {
    throw stageError("bucket_lookup", `Unknown bucket: ${bucketName}`, 400);
  }
  const sourceBucket = sourceLookup.bucket;
  const { bucket: libraryBucket } = getLibraryBucket(env);

  const object = await sourceBucket.get(canonicalKey);
  if (!object || !object.body) {
    throw stageError("get_object", "Object not found.", 404);
  }

  const { docId, basis, fieldsUsed, uploaded } = await computeDocId(bucketName, canonicalKey, {
    etag: object.etag,
    size: object.size,
    uploaded: (object as any)?.uploaded,
  });
  console.log("[LIBRARY_INGEST] start", {
    bucket: bucketName,
    key: canonicalKey,
    docId,
    etag: object.etag,
    size: object.size,
    uploaded: uploaded || (object as any)?.uploaded,
  });
  const extractedKey = buildExtractedPath(docId);
  const manifestKey = buildManifestPath(docId);
  const title = titleFromKey(canonicalKey);

  if (!skipCache) {
    try {
      const existing = typeof libraryBucket.head === "function"
        ? await libraryBucket.head(extractedKey)
        : await libraryBucket.get(extractedKey, { range: { offset: 0, length: 0 } as any });
      if (existing) {
        const manifest = await readManifest(libraryBucket, docId);
        await upsertLibraryIndexRecord(env, {
          docId,
          bucket: bucketName,
          key: canonicalKey,
          title,
          hashBasis: basis,
          hashFieldsUsed: fieldsUsed,
          normalizedTokens: tokensFromTitle(title),
          status: "ready",
          extractedKey,
          manifestKey,
          preview: manifest?.preview,
          size: object.size,
          etag: object.etag,
          uploaded: uploaded || (object as any)?.uploaded?.toISOString?.(),
        });
        console.log("[LIBRARY_INGEST] cache_hit", { docId, bucket: bucketName, key: canonicalKey, extractedKey });
        return { status: "cache_hit", docId, extractedKey, manifest };
      }
    } catch (err) {
      console.warn("[LIBRARY-INGEST] cache probe failed", { docId, extractedKey, error: err instanceof Error ? err.message : String(err) });
    }
  }

  const cachedManifest = await readManifest(libraryBucket, docId);
  if (cachedManifest && cachedManifest.method === "ocr" && !cachedManifest.pagesProcessed) {
    await upsertLibraryIndexRecord(env, {
      docId,
      bucket: bucketName,
      key: canonicalKey,
      title,
      hashBasis: basis,
      hashFieldsUsed: fieldsUsed,
      normalizedTokens: tokensFromTitle(title),
      status: "needs_browser_ocr",
      extractedKey,
      manifestKey,
      preview: cachedManifest.preview,
      size: object.size,
      etag: object.etag,
      uploaded: uploaded || (object as any)?.uploaded?.toISOString?.(),
    });
    console.log("[LIBRARY_INGEST] needs_browser_ocr", {
      docId,
      bucket: bucketName,
      key: canonicalKey,
      pageCount: cachedManifest?.pageCount,
    });
    return {
      status: "needs_browser_ocr",
      docId,
      bucket: bucketName,
      key: canonicalKey,
      pageCount: cachedManifest.pageCount,
      downloadUrl: buildLibraryDownloadUrl(bucketName, canonicalKey),
    };
  }

  const bytes = new Uint8Array(await object.arrayBuffer());
  let extraction;
  try {
    extraction = await extractEmbeddedPagesFromPdf(bytes, {
      sampleCount: PDF_SAMPLE_PAGE_TARGET,
      maxPages: mode === "embedded_only" ? LIBRARY_BATCH_PAGE_LIMIT : MAX_EXTRACT_PAGES,
      allowEarlyStop: true,
    });
  } catch (err) {
    throw stageError("extract_embedded", err instanceof Error ? err.message : "Embedded extraction failed.");
  }

  if (extraction.scanned || !extraction.pages.length) {
    const manifest: PdfManifest = {
      fileHash: docId,
      filename: title,
      method: "ocr",
      extractionMethod: "ocr",
      pagesProcessed: 0,
      pageCount: extraction.pageCount,
      createdAt: new Date().toISOString(),
      bucket: bucketName,
      key: canonicalKey,
      docId,
      hashBasis: basis,
      hashFieldsUsed: fieldsUsed,
      title,
    };
    try {
      await writeManifest(libraryBucket, manifest);
    } catch (err) {
      throw stageError("write_manifest", err instanceof Error ? err.message : "Failed to persist manifest.");
    }
    await upsertLibraryIndexRecord(env, {
      docId,
      bucket: bucketName,
      key: canonicalKey,
      title,
      hashBasis: basis,
      hashFieldsUsed: fieldsUsed,
      normalizedTokens: tokensFromTitle(title),
      status: "needs_browser_ocr",
      extractedKey,
      manifestKey,
      preview: "",
      size: object.size,
      etag: object.etag,
      uploaded: uploaded || (object as any)?.uploaded?.toISOString?.(),
    });
    return {
      status: "needs_browser_ocr",
      docId,
      bucket: bucketName,
      key: canonicalKey,
      pageCount: extraction.pageCount,
      downloadUrl: buildLibraryDownloadUrl(bucketName, canonicalKey),
    };
  }

  const normalized = normalizePages(extraction.pages);
  const finalText = normalized.text.slice(0, MAX_OCR_TEXT_LENGTH);
  const preview = finalText.slice(0, ONE_SHOT_PREVIEW_LIMIT);
  try {
    await libraryBucket.put(extractedKey, finalText, {
      httpMetadata: { contentType: "text/plain; charset=utf-8" },
    });
  } catch (err) {
    throw stageError("write_extracted", err instanceof Error ? err.message : "Failed to persist extracted text.");
  }
  const manifest: PdfManifest = {
    fileHash: docId,
    filename: title,
    method: "embedded",
    extractionMethod: "embedded",
    pagesProcessed: normalized.pagesProcessed,
    pageCount: extraction.pageCount,
    createdAt: new Date().toISOString(),
    preview,
    bucket: bucketName,
    key: canonicalKey,
    docId,
    hashBasis: basis,
    hashFieldsUsed: fieldsUsed,
    title,
  };
  try {
    await writeManifest(libraryBucket, manifest);
  } catch (err) {
    throw stageError("write_manifest", err instanceof Error ? err.message : "Failed to persist manifest.");
  }
  await upsertLibraryIndexRecord(env, {
    docId,
    bucket: bucketName,
    key: canonicalKey,
    title,
    hashBasis: basis,
    hashFieldsUsed: fieldsUsed,
    normalizedTokens: tokensFromTitle(title),
    status: "ready",
    extractedKey,
    manifestKey,
    preview,
    size: object.size,
    etag: object.etag,
    uploaded: uploaded || (object as any)?.uploaded?.toISOString?.(),
  });

  console.log("[LIBRARY_INGEST] ready", {
    docId,
    bucket: bucketName,
    key: canonicalKey,
    pages: extraction.pageCount,
    previewLen: preview.length,
    extractedKey,
  });
  return { status: "ok", docId, extractedKey, manifest, preview, pageCount: extraction.pageCount };
}

async function upsertLibraryIndexRecord(env: Env, record: LibraryIndexRecord) {
  const { bucket } = getLibraryBucket(env);
  const existing = await readLibraryIndex(bucket);
  const map = new Map<string, LibraryIndexRecord>();
  existing.forEach(rec => map.set(rec.docId, rec));

  const prior = map.get(record.docId);
  const normalizedTokens = record.normalizedTokens?.length
    ? record.normalizedTokens
    : prior?.normalizedTokens?.length
      ? prior.normalizedTokens
      : tokensFromTitle(record.title);

  const merged: LibraryIndexRecord = {
    ...(prior || {}),
    ...record,
    normalizedTokens,
    status: record.status || prior?.status,
    preview: normalizePreview(record.preview || prior?.preview),
    manifestKey: record.manifestKey || prior?.manifestKey || buildManifestPath(record.docId),
    extractedKey: record.extractedKey || prior?.extractedKey || buildExtractedPath(record.docId),
  };
  map.set(record.docId, merged);
  const nextList = Array.from(map.values());
  await writeLibraryIndex(bucket, nextList);
  await bucket.put(buildIndexKeyForDoc(record.docId), JSON.stringify(merged), {
    httpMetadata: { contentType: "application/json; charset=utf-8" },
  });
  return merged;
}

async function listPdfObjectsFromBucket(bucket: R2Bucket, opts: { limit?: number } = {}) {
  const results: R2Object[] = [];
  let cursor: string | undefined;
  const limit = opts.limit ?? 0;
  do {
    const page = await bucket.list({ cursor, limit: 1000 });
    for (const obj of page.objects ?? []) {
      if (isPdfKey(obj.key)) {
        results.push(obj);
        if (limit && results.length >= limit) {
          return results;
        }
      }
    }
    cursor = page.truncated ? page.cursor : undefined;
  } while (cursor);
  return results;
}

function buildLibraryDownloadUrl(bucket: string, key: string) {
  const base = `/api/file?bucket=${encodeURIComponent(bucket)}&key=${encodeURIComponent(key)}`;
  return base;
}

function buildRangesFromPages(pageIndices: number[]) {
  const sorted = Array.from(new Set(pageIndices.map(idx => Number(idx) + 1).filter(n => Number.isFinite(n)))).sort((a, b) => a - b);
  const ranges: Array<{ start: number; end: number }> = [];
  if (!sorted.length) return ranges;
  let start = sorted[0];
  let end = sorted[0];
  for (let i = 1; i < sorted.length; i += 1) {
    const current = sorted[i];
    if (current === end + 1) {
      end = current;
    } else {
      ranges.push({ start, end });
      start = current;
      end = current;
    }
  }
  ranges.push({ start, end });
  return ranges;
}

function mergeRanges(ranges: Array<{ start: number; end: number }>) {
  const cleaned = ranges
    .filter(range => Number.isFinite(range.start) && Number.isFinite(range.end))
    .map(range => (range.start <= range.end ? { start: range.start, end: range.end } : { start: range.end, end: range.start }))
    .sort((a, b) => a.start - b.start);
  if (!cleaned.length) return [];
  const merged: Array<{ start: number; end: number }> = [cleaned[0]];
  for (let i = 1; i < cleaned.length; i += 1) {
    const current = cleaned[i];
    const last = merged[merged.length - 1];
    if (current.start <= last.end + 1) {
      last.end = Math.max(last.end, current.end);
    } else {
      merged.push({ start: current.start, end: current.end });
    }
  }
  return merged;
}

function parseNormalizedPagesFromText(text: string) {
  const map = new Map<number, string>();
  const parts = (text || "").split(/---\s*Page\s+(\d+)\s*---/i);
  for (let i = 1; i < parts.length; i += 2) {
    const pageNumber = Number(parts[i]);
    const pageText = parts[i + 1] || "";
    if (Number.isFinite(pageNumber)) {
      map.set(pageNumber - 1, normalizePlainText(pageText));
    }
  }
  return map;
}

async function handleGenerateFile(req: Request, env: Env): Promise<Response> {
  const raw = await req.json().catch(() => null);
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return json({ error: "Send a JSON object body." }, 400);
  }

  const payload = raw as Partial<GenerateFileRequest>;
  const sanitizedMessages = sanitizeChatMessages(payload.messages);
  if (!sanitizedMessages.length) {
    return json({ error: "messages must include at least one non-empty entry." }, 400);
  }

  const agent = resolveAgent(typeof payload.agentId === "string" ? payload.agentId : undefined);
  const defaultAgent = resolveAgent(null);
  const desiredType = typeof payload.desiredFileType === "string" ? payload.desiredFileType : "txt";
  const ext = desiredType.replace(/^\./, "").toLowerCase() || "txt";
  const mime = mimeFromExtension(ext);
  const model = agent.model || env.DEFAULT_MODEL?.trim() || "gpt-5-mini";
  const fileSystemPrompt = `
You are OWEN, generating a single ${ext} file as output.
Rules:
Output ONLY the file contents, no explanations, no backticks, no surrounding markdown fences.
The user will download this text as a .${ext} file.
Do not include any commentary, only the raw file data.
`.trim();

  const userContent = sanitizedMessages
    .map(message => {
      const text = typeof message.content === "string"
        ? message.content
        : message.content.map(part => part.text).join("\n");
      return `${message.role.toUpperCase()}: ${text}`;
    })
    .join("\n\n");

  const upstream = await fetch(`${env.OPENAI_API_BASE}/chat/completions`, {
    method: "POST",
    headers: {
      authorization: `Bearer ${env.OPENAI_API_KEY}`,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model,
      messages: [
        { role: "system", content: fileSystemPrompt },
        { role: "user", content: userContent },
      ],
    }),
  });

  if (!upstream.ok) {
    const errText = await upstream.text().catch(() => "Unable to contact OpenAI.");
    return json({ error: errText }, upstream.status || 502);
  }

  const completion = await upstream.json().catch(() => null);
  const content = extractChatCompletionContent(completion);
  const targetBucketName = agent.defaultBuckets?.[0] || defaultAgent.defaultBuckets?.[0] || "OWEN_UPLOADS";
  const bucketLookup = lookupBucket(env, targetBucketName);
  if (!bucketLookup) {
    return json({ error: `Bucket ${targetBucketName} is not configured.` }, 500);
  }

  const key = `${Date.now()}-owen.${ext || "txt"}`;
  await bucketLookup.bucket.put(key, content, {
    httpMetadata: { contentType: mime },
  });

  return json({ bucket: bucketLookup.name, key });
}

async function handleExtract(req: Request, env: Env): Promise<Response> {
  const body = await req.json().catch(() => null);
  if (!body || typeof body !== "object") {
    return json({ error: "Send JSON { file_id }." }, 400);
  }
  console.log("extract endpoint: stripping attachment metadata");
  delete (body as any).attachments;
  delete (body as any).files;
  delete (body as any).fileRefs;
  const fileId = typeof (body as any).file_id === "string" ? (body as any).file_id.trim() : "";
  if (!fileId) {
    return json({ error: "Missing file_id." }, 400);
  }
  const filename = typeof (body as any).filename === "string" ? (body as any).filename : undefined;
  const model = typeof (body as any).model === "string" && (body as any).model.trim()
    ? (body as any).model.trim()
    : "gpt-5-mini";

  const prompt = [
    "You are a code interpreter assistant.",
    "Open the attached document, extract the text, and produce a concise summary.",
    "Respond strictly in JSON with keys:",
    `"summary" (<=200 words), "key_points" (array of short bullet strings),`,
    `"raw_excerpt" (first ~800 characters of raw text).`,
  ].join(" ");

  const payload = {
    model,
    text: { format: { type: "json_object" as const } },
    input: [
      {
        role: "user" as const,
        content: [
          { type: "input_file" as const, file_id: fileId },
          { type: "input_text" as const, text: prompt },
        ],
      },
    ],
    max_output_tokens: 800,
  };

  const upstream = await fetch(`${env.OPENAI_API_BASE}/responses`, {
    method: "POST",
    headers: {
      authorization: `Bearer ${env.OPENAI_API_KEY}`,
      "content-type": "application/json",
      "OpenAI-Beta": "assistants=v2",
    },
    body: JSON.stringify(payload),
  });

  if (!upstream.ok) {
    const errText = await upstream.text().catch(() => "Unable to contact OpenAI.");
    return json({ error: errText }, upstream.status || 502);
  }

  const data = await upstream.json().catch(() => ({}));
  const text = extractOutputText(data).trim();
  let parsed: any = {};
  try {
    parsed = text ? JSON.parse(text) : {};
  } catch {
    parsed.summary = text;
  }

  // Some models occasionally nest the JSON as a string; unwrap it when detected.
  if (typeof parsed.summary === "string") {
    try {
      const nested = JSON.parse(parsed.summary);
      if (nested && typeof nested === "object") {
        if (typeof nested.summary === "string") parsed.summary = nested.summary;
        if (Array.isArray(nested.key_points)) parsed.key_points = nested.key_points;
        if (typeof nested.raw_excerpt === "string") parsed.raw_excerpt = nested.raw_excerpt;
        if (typeof nested.filename === "string" && !parsed.filename) parsed.filename = nested.filename;
      }
    } catch {}
  }

  return json({
    summary: typeof parsed.summary === "string" ? parsed.summary : undefined,
    key_points: Array.isArray(parsed.key_points) ? parsed.key_points.filter((point: unknown): point is string => typeof point === "string") : [],
    raw_excerpt: typeof parsed.raw_excerpt === "string" ? parsed.raw_excerpt : undefined,
    filename,
  });
}

async function performDocumentTextExtraction(env: Env, opts: { fileId: string; visionFileId?: string | null; filename: string; mimeType?: string | null }) {
  const { fileId, visionFileId, filename, mimeType } = opts;
  const kind = inferIngestKind(filename, mimeType);
  // Prefer direct vision OCR for images when available.
  if (kind === "image" && visionFileId) {
    try {
      const visionText = await requestVisionTextExtraction(env, visionFileId, filename);
      if (visionText.trim()) {
        console.log("Vision OCR succeeded", { filename, textLength: visionText.length });
        return `[OCR]\n${visionText.trim()}`;
      }
    } catch (err) {
      console.warn("Vision OCR failed, falling back to doc extraction", { filename, error: err instanceof Error ? err.message : String(err) });
    }
  }

  const failures: Array<{ model: string; message: string }> = [];
  try {
    console.log("[PDF] Attempting OCR extraction via /responses", { filename, kind });
    const text = await requestDocumentTextExtraction(env, fileId, filename);
    let cleaned = (text || "").trim();
    const failurePhrases = [
      "unable to open files directly",
      "can't read the pdf",
      "cant read the pdf",
      "cannot read the pdf",
      "please re-upload the file",
      "could not read the file",
    ];
    if (failurePhrases.some(phrase => cleaned.toLowerCase().includes(phrase))) {
      cleaned = "";
    }
    if (cleaned) {
      // If PDF extraction is suspiciously short, attempt a vision OCR fallback when possible.
      if (kind === "pdf" && cleaned.length < 500 && visionFileId) {
        try {
          console.log("[PDF] Extraction is short; trying vision OCR fallback", { filename });
          const visionText = await requestVisionTextExtraction(env, visionFileId, filename);
          if (visionText.trim()) {
            return `[OCR]\n${visionText.trim()}`;
          }
        } catch (visionErr) {
          const msg = visionErr instanceof Error ? visionErr.message : String(visionErr);
          failures.push({ model: "gpt-4o(vision)", message: msg });
          console.warn("[PDF] Vision OCR fallback failed", { filename, message: msg });
        }
      }
      console.log("[PDF] OCR extraction succeeded", { model: "gpt-4o", textLength: cleaned.length });
      return cleaned.slice(0, MAX_OCR_TEXT_LENGTH);
    } else if (!cleaned && visionFileId) {
      try {
        console.log("[PDF] Primary OCR empty; trying vision OCR fallback", { filename });
        const visionText = await requestVisionTextExtraction(env, visionFileId, filename);
        if (visionText.trim()) {
          return `[OCR]\n${visionText.trim()}`;
        }
      } catch (visionErr) {
        const msg = visionErr instanceof Error ? visionErr.message : String(visionErr);
        failures.push({ model: "gpt-4o(vision)", message: msg });
        console.warn("[PDF] Vision OCR fallback after empty result failed", { filename, message: msg });
      }
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    failures.push({ model: "gpt-4o", message });
    console.warn("OCR extraction failed", { model: "gpt-4o", message });
  }
  const detail = failures.length
    ? failures.map(entry => `[${entry.model}] ${entry.message}`).join("; ")
    : "No OCR models available.";
  throw new Error(`OCR request failed: ${detail}`);
}

type IngestKind = "text" | "pdf" | "image" | "office" | "other";

function inferIngestKind(filename: string, mimeType?: string | null): IngestKind {
  const name = (filename || "").toLowerCase();
  const mime = (mimeType || "").toLowerCase();
  if (mime.startsWith("image/") || isLikelyImageFilename(name)) return "image";
  if (mime.includes("pdf") || isLikelyPdfFilename(name)) return "pdf";
  if (mime.startsWith("text/") || isTextKeyByExtension(name)) return "text";
  if (mime.includes("msword") || mime.includes("officedocument") || mime.includes("powerpoint")) return "office";
  return "other";
}

type OcrPrimitiveOptions = { label?: string; maxOutputTokens?: number };

async function callResponsesOcr(
  env: Env,
  content: ResponsesInputContent[],
  opts: OcrPrimitiveOptions = {},
): Promise<string> {
  const base = env.OPENAI_API_BASE?.replace(/\/$/, "") || "https://api.openai.com/v1";
  const maxTokens = Math.min(OCR_MAX_OUTPUT_TOKENS, Math.max(200, opts.maxOutputTokens ?? OCR_MAX_OUTPUT_TOKENS));
  const payload = {
    model: resolveModelId("gpt-4o", env),
    input: [
      {
        role: "user" as const,
        content,
      },
    ],
    max_output_tokens: maxTokens,
  };

  console.log("[OCR] Sending /responses request", {
    label: opts.label || "ocr",
    contentTypes: content.map(part => part.type),
  });

  const resp = await retryOpenAI(
    () =>
      fetch(`${base}/responses`, {
        method: "POST",
        headers: {
          authorization: `Bearer ${env.OPENAI_API_KEY}`,
          "content-type": "application/json",
        },
        body: JSON.stringify(payload),
      }),
    opts.label || "ocr",
  );

  const data = await safeJson(resp);
  if (!resp.ok) {
    const message = data?.error?.message || data?.message || resp.statusText || "OCR request failed.";
    console.error("[OCR] OpenAI error", { label: opts.label || "ocr", status: resp.status, body: data });
    throw new Error(message);
  }

  const raw = extractOutputText(data).trim();
  return raw.slice(0, MAX_OCR_TEXT_LENGTH);
}

async function requestVisionTextExtraction(env: Env, visionFileId: string, filename: string) {
  const prompt = [
    "You are an OCR service helping Owen answer questions.",
    "Read the attached image carefully and transcribe all visible text.",
    "Preserve layout where possible: headings, bullets, table cells.",
    "If multi-page, prepend each page with markers like [Page 1], [Page 2].",
    "Use ? for uncertain characters; do not invent text.",
    'Respond strictly as JSON: { "pages": [ { "page": "<label or number>", "text": "<clean text>" } ] } and include at least one entry.',
  ].join(" ");

  const raw = await callResponsesOcr(
    env,
    [
      { type: "input_text", text: prompt },
      { type: "input_image", file_id: visionFileId, detail: "high" as const },
    ],
    { label: `vision-file:${filename}` },
  );
  if (!raw) return "";

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed?.pages)) {
      const formatted = formatExtractedPages(parsed.pages);
      if (formatted) return formatted;
    }
    if (typeof parsed?.text === "string" && parsed.text.trim()) {
      return parsed.text.trim().slice(0, MAX_OCR_TEXT_LENGTH);
    }
  } catch {
    // fall through to raw handling
  }
  return raw.slice(0, MAX_OCR_TEXT_LENGTH);
}

async function requestDocumentTextExtraction(env: Env, fileId: string, filename: string) {
  const prompt = [
    `You are a document ingestion (text extraction + OCR) service helping OWEN answer questions.`,
    `Open the attached file (${filename || "Document"}). The file may be text-based or scanned.`,
    `For text-based PDFs: extract embedded text. For scanned PDFs/images: render pages (if needed) and OCR.`,
    `Extract full text in reading order. Preserve headings, bullet points, and table cell contents.`,
    `If the document has pages or slides, prepend each section with markers like [Page 1] or [Slide 2].`,
    `Do NOT summarize. Do NOT fabricate text. Use ? for uncertain characters.`,
    `Respond strictly as JSON: { "pages": [ { "page": "<label or number>", "text": "<clean text>" } ] } and include at least one entry.`,
  ].join(" ");

  const raw = await callResponsesOcr(
    env,
    [
      {
        type: "input_text",
        text: prompt,
      },
      {
        type: "input_file",
        file_id: fileId,
      },
    ],
    { label: `doc-file:${filename}` },
  );
  if (!raw) return "";

  let parsed: any;
  try {
    parsed = JSON.parse(raw);
  } catch {
    parsed = null;
  }
  if (Array.isArray(parsed?.pages)) {
    const formatted = formatExtractedPages(parsed.pages);
    if (formatted) return formatted;
  }
  if (typeof parsed?.text === "string" && parsed.text.trim()) {
    return parsed.text.trim().slice(0, MAX_OCR_TEXT_LENGTH);
  }
  return raw.slice(0, MAX_OCR_TEXT_LENGTH);
}

function formatExtractedPages(pages: any[]): string {
  let buffer = "";
  for (const entry of pages) {
    if (!entry || typeof entry !== "object") continue;
    const rawNumber = "page" in entry ? entry.page : entry.page_number ?? entry.number;
    const pageNumber = typeof rawNumber === "number" ? rawNumber : Number(rawNumber);
    const label =
      typeof rawNumber === "string" && rawNumber.trim()
        ? rawNumber.trim()
        : Number.isFinite(pageNumber)
          ? `Page ${pageNumber}`
          : "Page";
    const text = typeof entry.text === "string" ? entry.text.trim() : "";
    if (!text) continue;
    const block = `[${label}] ${text}\n\n`;
    buffer += block;
    if (buffer.length >= MAX_OCR_TEXT_LENGTH) {
      return buffer.slice(0, MAX_OCR_TEXT_LENGTH);
    }
  }
  const trimmed = buffer.trim();
  return trimmed.slice(0, MAX_OCR_TEXT_LENGTH);
}

const MIN_PDF_TEXT_LENGTH_FOR_EMBEDDED = 200;
const OCR_IMAGE_DETAIL: InputImageDetail = "high";
const OCR_PDF_RENDER_MAX_DIMENSION = 1600;
const OCR_PDF_RENDER_BASE_SCALE = 2;
const OCR_BATCH_IMAGE_COUNT = 3;

type TranscriptIngestResult = { text: string; source: "ocr" | "original" };
type ImageSignature = "png" | "jpg" | "unknown";
type PreparedOcrImage =
  | {
      ok: true;
      dataUrl: string;
      mimeType: string;
      signature: ImageSignature;
      head: string;
      byteLength: number;
    }
  | {
      ok: false;
      mimeType: string;
      signature: ImageSignature;
      head: string;
      byteLength: number;
    };

function isLikelyPdfFilename(name: string) {
  return /\.pdf$/i.test(name || "");
}

function isLikelyImageFilename(name: string) {
  return /\.(png|jpe?g|gif|webp|bmp|tiff?|heic)$/i.test(name || "");
}

function isPdfMimeType(mimeType: string) {
  return (mimeType || "").toLowerCase().includes("pdf");
}

function isImageMimeType(mimeType: string) {
  return (mimeType || "").toLowerCase().startsWith("image/");
}

function guessImageMimeTypeFromFilename(name: string): string {
  const lowered = (name || "").toLowerCase();
  if (lowered.endsWith(".png")) return "image/png";
  if (lowered.endsWith(".jpg") || lowered.endsWith(".jpeg")) return "image/jpeg";
  if (lowered.endsWith(".webp")) return "image/webp";
  if (lowered.endsWith(".gif")) return "image/gif";
  if (lowered.endsWith(".bmp")) return "image/bmp";
  if (lowered.endsWith(".tif") || lowered.endsWith(".tiff")) return "image/tiff";
  if (lowered.endsWith(".heic")) return "image/heic";
  return "";
}

function bytesToBase64(bytes: Uint8Array): string {
  const chunkSize = 0x2000;
  let binary = "";
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function detectImageSignature(bytes: Uint8Array): ImageSignature {
  if (bytes.length >= 4 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    return "png";
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return "jpg";
  }
  return "unknown";
}

function normalizeOcrImageMime(mimeType: string, signature: ImageSignature): string {
  if (signature === "png") return "image/png";
  if (signature === "jpg") return "image/jpeg";
  if ((mimeType || "").toLowerCase().startsWith("image/")) return mimeType;
  return "application/octet-stream";
}

function prepareOcrImageInput(bytes: Uint8Array, mimeType: string): PreparedOcrImage {
  const signature = detectImageSignature(bytes);
  const normalizedMime = normalizeOcrImageMime(mimeType, signature);
  const base64 = bytesToBase64(bytes);
  const head = base64.slice(0, 12);
  if (signature === "unknown") {
    return {
      ok: false,
      mimeType: normalizedMime,
      signature,
      head,
      byteLength: bytes.length,
    };
  }
  const dataUrl = `data:${normalizedMime};base64,${base64}`;
  return {
    ok: true,
    dataUrl,
    mimeType: normalizedMime,
    signature,
    head,
    byteLength: bytes.length,
  };
}

function inspectDataUrlMeta(dataUrl: string): { mimeType: string; head: string; signature: ImageSignature; byteLength?: number } {
  const match = /^data:([^;]+);base64,([a-zA-Z0-9+/=]+)/.exec(dataUrl) || [];
  const mimeType = match[1] || "";
  const b64 = match[2] || "";
  const head = b64.slice(0, 12);
  let signature: ImageSignature = "unknown";
  try {
    const sliceLen = Math.min(Math.max(8, head.length), b64.length);
    const padded = b64.slice(0, sliceLen);
    const padLength = padded.length % 4 === 0 ? 0 : 4 - (padded.length % 4);
    const decoded = atob(padded.padEnd(padded.length + padLength, "="));
    const probe = new Uint8Array(decoded.length);
    for (let i = 0; i < decoded.length; i += 1) {
      probe[i] = decoded.charCodeAt(i);
    }
    signature = detectImageSignature(probe);
  } catch {
    signature = "unknown";
  }
  const byteLength = b64 ? Math.floor((b64.length * 3) / 4) : undefined;
  return { mimeType, head, signature, byteLength };
}

function logOcrInputDebug(meta: { key?: string; head?: string; mimeType?: string; signature?: ImageSignature; byteLength?: number }) {
  console.log("[OCR_INPUT]", {
    key: meta.key || "unknown",
    head: (meta.head || "").slice(0, 12),
    mime: meta.mimeType || "unknown",
    sig: meta.signature || "unknown",
    len: meta.byteLength ?? 0,
  });
}

function invalidImagePayload(extra: Record<string, unknown> = {}) {
  return {
    ok: false,
    error: { code: "invalid_image_bytes" },
    answer: "OCR input was not a valid image.",
    ...extra,
  };
}

function normalizeExtractedText(value: string): string {
  let text = normalizePlainText(value);
  if (!text) return text;

  // Ensure page markers are separated and spaced
  text = text.replace(/([a-z])(?=page\s*\d+)/gi, "$1 ");
  text = text.replace(/(---\s*page\s*\d+\s*---)(?=\S)/gi, "$1\n");
  text = text.replace(/(page\s*\d+)(?=\S)/gi, "$1 ");

  // Remove obvious OCR junk lines
  const junkLines = new Set(["dh", "r", "met"]);
  const lines = text.split("\n");
  const cleanedLines: string[] = [];
  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      cleanedLines.push("");
      return;
    }
    const lower = trimmed.toLowerCase();
    if (junkLines.has(lower)) return;
    if (/^kcu[-\s]?com/i.test(trimmed)) return;
    if (/^dr\./i.test(trimmed) && trimmed === trimmed.toUpperCase()) return;
    cleanedLines.push(line);
  });

  text = cleanedLines.join("\n");

  // Repair hyphenation artifacts lightly
  text = text.replace(/-\s*\n\s*/g, "");
  text = text.replace(/\s{3,}/g, " ");
  text = text.replace(/\n{3,}/g, "\n\n");
  return normalizePlainText(text);
}

async function ingestBinaryForTranscript(
  env: Env,
  {
    bytes,
    filename,
    mimeType,
    sourceKey,
  }: {
    bytes: Uint8Array;
    filename: string;
    mimeType: string;
    sourceKey?: string;
  },
): Promise<TranscriptIngestResult | null> {
  const name = filename || "document";
  const mime = mimeType || "";

  if (isTextKeyByExtension(name) || isTextContentType(mime)) {
    try {
      const decoded = new TextDecoder().decode(bytes);
      const normalized = normalizeExtractedText(decoded);
      if (!normalized) return null;
      return { text: normalized.slice(0, MAX_OCR_TEXT_LENGTH), source: "original" };
    } catch {
      // fall through
    }
  }

  if (isLikelyPdfFilename(name) || isPdfMimeType(mime)) {
    if (PDFJS_AVAILABLE) {
      try {
        const embedded = await extractEmbeddedTextFromPdf(bytes);
        if (embedded.length >= MIN_PDF_TEXT_LENGTH_FOR_EMBEDDED) {
          return { text: embedded.slice(0, MAX_OCR_TEXT_LENGTH), source: "original" };
        }
      } catch (err) {
        console.warn("Embedded PDF extraction failed; continuing to OCR fallback", {
          filename: name,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    } else {
      console.log("PDF.js unavailable in Worker; skipping embedded extraction.", { filename: name });
    }
    // If pdfjs is not available, return null to let the legacy CI extraction path run.
    return null;
  }

  if (isLikelyImageFilename(name) || isImageMimeType(mime)) {
    try {
      const inferredMime = isImageMimeType(mime) ? mime : guessImageMimeTypeFromFilename(name) || "image/png";
      const prepared = prepareOcrImageInput(bytes, inferredMime);
      if (!prepared.ok) {
        console.warn("Image OCR ingestion blocked due to invalid bytes", {
          filename: name,
          key: sourceKey,
          head: prepared.head,
          mimeType: prepared.mimeType,
        });
        return null;
      }
      const extracted = await requestVisionOcrFromImages(
        env,
        [
          {
            label: "Image OCR",
            dataUrl: prepared.dataUrl,
            sourceKey: sourceKey || name,
            mimeType: prepared.mimeType,
            signature: prepared.signature,
            byteLength: prepared.byteLength,
            head: prepared.head,
          },
        ],
        name,
      );
      if (!extracted) return null;
      const tagged = `[OCR]\n${extracted}`.slice(0, MAX_OCR_TEXT_LENGTH);
      return { text: tagged, source: "ocr" };
    } catch (err) {
      console.warn("Image OCR ingestion failed; allowing fallback", {
        filename: name,
        error: err instanceof Error ? err.message : String(err),
      });
      return null;
    }
  }

  return null;
}

function buildSamplePageNumbers(pageCount: number, target = PDF_SAMPLE_PAGE_TARGET): number[] {
  const total = Math.max(1, Math.min(pageCount, target));
  if (total >= pageCount) {
    return Array.from({ length: pageCount }, (_, i) => i + 1);
  }
  const pages = new Set<number>();
  pages.add(1);
  pages.add(pageCount);
  for (let i = 0; pages.size < total; i += 1) {
    const ratio = (i + 1) / (total + 1);
    const page = Math.max(1, Math.min(pageCount, Math.round(pageCount * ratio)));
    pages.add(page);
  }
  return Array.from(pages).sort((a, b) => a - b);
}

async function extractPdfPageText(doc: any, pageNum: number): Promise<string> {
  const page = await doc.getPage(pageNum);
  const content = await page.getTextContent();
  let buffer = "";
  for (const item of content.items as any[]) {
    const str = typeof item?.str === "string" ? item.str : "";
    if (!str) continue;
    buffer += str;
    buffer += item?.hasEOL ? "\n" : " ";
  }
  try {
    await page.cleanup?.();
  } catch {}
  return normalizePlainText(buffer);
}

async function extractEmbeddedPagesFromPdf(
  pdfBytes: Uint8Array,
  opts: { sampleOnly?: boolean; allowEarlyStop?: boolean; maxPages?: number; sampleCount?: number } = {},
): Promise<{ pages: PageText[]; pageCount: number; sampleTextLength: number; sampledPagesWithText: number; scanned: boolean }> {
  if (!PDFJS_AVAILABLE) {
    console.warn("PDF.js unavailable; skipping embedded text extraction.");
    return { pages: [], pageCount: 0, sampleTextLength: 0, sampledPagesWithText: 0, scanned: true };
  }
  const task = pdfjs.getDocument({
    data: pdfBytes,
    disableWorker: true,
  });
  const doc = await task.promise;
  const pageCount = doc?.numPages || 0;
  if (!pageCount) {
    return { pages: [], pageCount: 0, sampleTextLength: 0, sampledPagesWithText: 0, scanned: true };
  }
  const maxPages = Math.max(1, Math.min(opts.maxPages ?? MAX_EXTRACT_PAGES, pageCount || MAX_EXTRACT_PAGES));
  const samplePages = buildSamplePageNumbers(pageCount, opts.sampleCount ?? PDF_SAMPLE_PAGE_TARGET);
  const seen = new Set<number>();
  const pages: PageText[] = [];
  let sampleTextLength = 0;
  let sampledPagesWithText = 0;

  try {
    for (const pageNum of samplePages) {
      const normalized = await extractPdfPageText(doc, pageNum);
      if (normalized) {
        pages.push({ pageIndex: pageNum - 1, text: normalized });
        sampleTextLength += normalized.length;
        sampledPagesWithText += 1;
      }
      seen.add(pageNum);
    }

    const treatAsEmbedded = sampleTextLength >= MIN_EMBEDDED_PDF_CHARS || sampledPagesWithText >= 3;
    if (opts.sampleOnly || (!treatAsEmbedded && opts.allowEarlyStop !== false)) {
      return { pages, pageCount, sampleTextLength, sampledPagesWithText, scanned: !treatAsEmbedded };
    }

    for (let pageNum = 1; pageNum <= maxPages; pageNum += 1) {
      if (seen.has(pageNum)) continue;
      const normalized = await extractPdfPageText(doc, pageNum);
      if (normalized) {
        pages.push({ pageIndex: pageNum - 1, text: normalized });
      }
    }
    pages.sort((a, b) => a.pageIndex - b.pageIndex);

    return { pages, pageCount, sampleTextLength, sampledPagesWithText, scanned: false };
  } finally {
    try {
      await doc.cleanup?.();
      await doc.destroy?.();
    } catch {}
  }
}

async function extractEmbeddedTextFromPdf(pdfBytes: Uint8Array): Promise<string> {
  const result = await extractEmbeddedPagesFromPdf(pdfBytes, { allowEarlyStop: false, sampleCount: PDF_SAMPLE_PAGE_TARGET });
  const normalized = normalizePages(result.pages);
  return normalized.text;
}

async function renderPdfPageToPngDataUrl(page: any): Promise<string> {
  if (typeof OffscreenCanvas === "undefined") {
    throw new Error("OffscreenCanvas is unavailable in this runtime; skipping PDF render OCR path.");
  }
  const baseViewport = page.getViewport({ scale: 1 });
  const baseMax = Math.max(baseViewport.width, baseViewport.height);
  let scale = OCR_PDF_RENDER_BASE_SCALE;
  if (baseMax * scale > OCR_PDF_RENDER_MAX_DIMENSION) {
    scale = OCR_PDF_RENDER_MAX_DIMENSION / baseMax;
  }
  if (!Number.isFinite(scale) || scale <= 0) scale = 1;

  const viewport = page.getViewport({ scale });
  const width = Math.max(1, Math.ceil(viewport.width));
  const height = Math.max(1, Math.ceil(viewport.height));

  const canvas = new OffscreenCanvas(width, height);
  const ctx = canvas.getContext("2d", { alpha: false });
  if (!ctx) {
    throw new Error("OffscreenCanvas 2D context unavailable (required for PDF OCR rendering).");
  }

  (ctx as any).fillStyle = "white";
  (ctx as any).fillRect(0, 0, width, height);

  const renderTask = page.render({ canvasContext: ctx, viewport });
  await renderTask.promise;

  const blob = await canvas.convertToBlob({ type: "image/png" });
  const ab = await blob.arrayBuffer();
  const b64 = bytesToBase64(new Uint8Array(ab));
  return `data:image/png;base64,${b64}`;
}

async function ocrPdfBytesWithVision(env: Env, pdfBytes: Uint8Array, filename: string): Promise<string> {
  console.warn("[PDF] OCR render path is disabled in Worker runtime.");
  return "";
}

async function requestVisionOcrFromImages(
  env: Env,
  pages: Array<{
    label: string;
    dataUrl: string;
    sourceKey?: string;
    mimeType?: string;
    signature?: ImageSignature;
    byteLength?: number;
    head?: string;
  }>,
  filename: string,
): Promise<string> {
  const prompt = [
    "You are an OCR transcription service.",
    "Extract all readable text verbatim from EACH provided image.",
    "Preserve headings, bullets, tables, and line breaks as best as possible.",
    "If a character is unclear, replace it with '?'.",
    `Return strictly JSON: { "pages": [ { "page": "<label>", "text": "<text>" } ] }.`,
    "Use the label provided immediately before each image as the page value.",
  ].join(" ");

  const content: ResponsesInputContent[] = [{ type: "input_text", text: prompt }];
  const validatedPages: Array<{
    label: string;
    dataUrl: string;
    sourceKey?: string;
    mimeType: string;
    signature: ImageSignature;
    byteLength?: number;
    head?: string;
  }> = [];

  for (const page of pages) {
    const inspected = inspectDataUrlMeta(page.dataUrl);
    const signature = page.signature || inspected.signature;
    const mimeType = page.mimeType || inspected.mimeType || "application/octet-stream";
    const head = page.head || inspected.head;
    const byteLength = page.byteLength ?? inspected.byteLength;
    if (signature === "unknown") {
      console.warn("[OCR-PREP] Invalid image signature for OCR", {
        key: page.sourceKey || page.label || filename,
        head,
        mimeType,
        signature,
      });
      return "";
    }
    validatedPages.push({ ...page, mimeType, signature, byteLength, head });
    content.push({ type: "input_text", text: `LABEL: ${page.label}` });
    content.push({ type: "input_image", image_url: page.dataUrl, detail: OCR_IMAGE_DETAIL });
  }

  validatedPages.forEach(page =>
    logOcrInputDebug({
      key: page.sourceKey || page.label || filename,
      head: page.head,
      mimeType: page.mimeType,
      signature: page.signature,
      byteLength: page.byteLength,
    }),
  );

  const raw = await callResponsesOcr(env, content, { label: `vision-batch:${filename}` });
  if (!raw) return "";

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed?.pages)) {
      return formatExtractedPages(parsed.pages);
    }
    if (typeof parsed?.text === "string" && parsed.text.trim()) {
      return parsed.text.trim().slice(0, MAX_OCR_TEXT_LENGTH);
    }
  } catch {
    // fall through
  }

  return raw.slice(0, MAX_OCR_TEXT_LENGTH);
}

async function requestVisionOcrFromFileId(
  env: Env,
  visionFileId: string,
  label: string,
): Promise<string> {
  const prompt = [
    "You are an OCR transcription service.",
    "Extract all readable text verbatim from the attached image/PDF file.",
    "Preserve headings, bullets, tables, and line breaks as best as possible.",
    "If a character is unclear, replace it with '?'.",
    "Return strictly JSON: { \"pages\": [ { \"page\": \"<label>\", \"text\": \"<text>\" } ] }.",
  ].join(" ");

  const raw = await callResponsesOcr(
    env,
    [
      { type: "input_text" as const, text: prompt },
      { type: "input_text" as const, text: `LABEL: ${label || "Page"}` },
      { type: "input_image" as const, file_id: visionFileId, detail: "high" as InputImageDetail },
    ],
    { label: `vision-file:${label}` },
  );
  if (!raw) return "";
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed?.pages)) {
      return formatExtractedPages(parsed.pages);
    }
    if (typeof parsed?.text === "string" && parsed.text.trim()) {
      return parsed.text.trim().slice(0, MAX_OCR_TEXT_LENGTH);
    }
  } catch {
    // fall through
  }
  return raw.slice(0, MAX_OCR_TEXT_LENGTH);
}

function shouldAttemptTextExtraction(filename: string, mime: string) {
  const loweredName = (filename || "").toLowerCase();
  const loweredMime = (mime || "").toLowerCase();
  const docExts = [".pdf", ".doc", ".docx", ".ppt", ".pptx", ".pptm", ".docm", ".rtf", ".odt", ".odp"];
  const imageExts = [".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp", ".heic"];
  if (docExts.some(ext => loweredName.endsWith(ext))) return true;
  if (imageExts.some(ext => loweredName.endsWith(ext))) return true;
  if (loweredMime.includes("pdf")) return true;
  if (loweredMime.startsWith("image/")) return true;
  if (loweredMime.includes("msword") || loweredMime.includes("officedocument") || loweredMime.includes("powerpoint")) {
    return true;
  }
  return false;
}

function resolveBucket(env: Env, choice: FormDataEntryValue | null) {
  const requested = typeof choice === "string" ? choice : null;
  const defaultBinding = BUCKET_BINDINGS[DEFAULT_BUCKET];
  const lookup =
    lookupBucket(env, requested) ??
    lookupBucket(env, defaultBinding) ??
    lookupBucket(env, DEFAULT_BUCKET);
  if (!lookup) {
    throw new Error("No matching R2 bucket binding configured.");
  }
  return lookup;
}

function resolveUploadPrefix(choice: FormDataEntryValue | null): string {
  const value = typeof choice === "string" ? choice : "";
  if (!value) return "";
  switch (value) {
    case "anki_decks":
      return "Anki Decks/";
    case "study_guides":
      return "Study Guides/";
    case "library":
      return "library/";
    default:
      return "";
  }
}

function getBucketByName(env: Env, name: string) {
  const lookup = lookupBucket(env, name);
  if (!lookup) throw new Error(`Unknown bucket ${name}`);
  return lookup.bucket;
}

async function handleFile(req: Request, env: Env): Promise<Response> {
  const url = new URL(req.url);
  const searchKey = url.searchParams.get("key");
  if (!searchKey) return json({ error: "Missing key parameter" }, 400);
  const requestedBucket = url.searchParams.get("bucket");
  const candidates = requestedBucket
    ? [requestedBucket]
    : Object.keys(BUCKET_BINDINGS);

  let object: R2ObjectBody | null = null;
  let bucketUsed: string | null = null;
  let matchedKey = searchKey;

  for (const name of candidates) {
    try {
      const bucket = getBucketByName(env, name);
      const match = await findObjectInBucket(bucket, searchKey);
      if (match) {
        object = match.object;
        bucketUsed = name;
        matchedKey = match.key;
        break;
      }
    } catch {
      continue;
    }
  }
  if (!object || !object.body) {
    return new Response("Not found", { status: 404, headers: CORS_HEADERS });
  }
  const headers = new Headers(CORS_HEADERS);
  const contentType = object.httpMetadata?.contentType || "application/octet-stream";
  headers.set("content-type", contentType);
  const filenameParam = url.searchParams.get("filename");
  const filename = sanitizeFilename(filenameParam || matchedKey.split("/").pop() || "download.bin");
  headers.set("content-disposition", `attachment; filename="${filename}"`);
  if (bucketUsed) headers.set("x-owen-bucket", bucketUsed);
  return new Response(object.body, { headers });
}

async function handleDownload(req: Request, env: Env): Promise<Response> {
  const url = new URL(req.url);
  const bucketName = url.searchParams.get("bucket");
  const key = url.searchParams.get("key");
  if (!bucketName || !key) {
    return new Response("Missing bucket or key", { status: 400, headers: CORS_HEADERS });
  }
  const lookup = lookupBucket(env, bucketName);
  if (!lookup) {
    return new Response("Unknown bucket", { status: 400, headers: CORS_HEADERS });
  }
  const object = await lookup.bucket.get(key);
  if (!object || !object.body) {
    return new Response("Not found", { status: 404, headers: CORS_HEADERS });
  }
  return new Response(object.body, {
    headers: {
      ...CORS_HEADERS,
      "Content-Type": object.httpMetadata?.contentType || "application/octet-stream",
      "Content-Disposition": `attachment; filename="${key}"`,
    },
  });
}

function json(obj: unknown, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", ...CORS_HEADERS, ...NO_STORE_HEADERS },
  });
}

function jsonNoStore(obj: unknown, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", ...CORS_HEADERS, ...NO_STORE_HEADERS },
  });
}

function mimeFromExtension(ext: string) {
  switch (ext) {
    case "md":
      return "text/markdown";
    case "json":
      return "application/json";
    case "csv":
      return "text/csv";
    case "txt":
    default:
      return "text/plain";
  }
}

function extractChatCompletionContent(payload: any) {
  const message = payload?.choices?.[0]?.message;
  if (!message) return "";
  const content = message.content;
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content
      .map(part => {
        if (typeof part === "string") return part;
        if (typeof part?.text === "string") return part.text;
        return "";
      })
      .filter(Boolean)
      .join("");
  }
  return "";
}

function encodeSSE({ event, data }: { event: string; data: string }) {
  return enc.encode(`event: ${event}\ndata: ${data}\n\n`);
}

function sanitizeKey(input: string | null, fallback: string) {
  if (!input) return fallback;
  const safe = input
    .trim()
    .replace(/\\/g, "/")
    .replace(/[^a-zA-Z0-9._/-]+/g, "-")
    .replace(/\/+/g, "/")
    .replace(/^\.+/, "");
  return safe || fallback;
}

function sanitizeFilename(name: string) {
  return name.replace(/[^\w.\-]+/g, "_");
}

async function findObjectInBucket(bucket: R2Bucket, candidateKey: string) {
  const attempts = new Set<string>();
  const trimmed = candidateKey.trim();
  if (trimmed) attempts.add(trimmed);
  const sanitized = sanitizeKey(trimmed, trimmed);
  if (sanitized) attempts.add(sanitized);

  for (const key of attempts) {
    try {
      const obj = await bucket.get(key);
      if (obj && obj.body) return { object: obj, key };
    } catch {}
  }

  for (const prefix of attempts) {
    if (!prefix) continue;
    try {
      const list = await bucket.list({ prefix });
      const match = list.objects?.[0];
      if (match) {
        const obj = await bucket.get(match.key);
        if (obj && obj.body) return { object: obj, key: match.key };
      }
    } catch {}
  }

  try {
    const list = await bucket.list();
    const match = list.objects?.find(obj =>
      Array.from(attempts).some(val => val && obj.key.endsWith(val)),
    );
    if (match) {
      const obj = await bucket.get(match.key);
      if (obj && obj.body) return { object: obj, key: match.key };
    }
  } catch {}

  return null;
}

function lookupBucket(env: Env, identifier?: string | null) {
  if (!identifier) return null;
  const trimmed = identifier.trim();
  if (!trimmed) return null;
  const envRecord = env as unknown as Record<string, unknown>;

  const direct = envRecord[trimmed];
  if (isR2BucketLike(direct)) {
    return { bucket: direct as R2Bucket, name: trimmed };
  }

  const upper = trimmed.toUpperCase();
  if (upper !== trimmed) {
    const upperMatch = envRecord[upper];
    if (isR2BucketLike(upperMatch)) {
      return { bucket: upperMatch as R2Bucket, name: upper };
    }
  }

  const lower = trimmed.toLowerCase();
  if (lower in BUCKET_BINDINGS) {
    const bindingName = BUCKET_BINDINGS[lower as keyof typeof BUCKET_BINDINGS];
    const bindingBucket = envRecord[bindingName];
    if (isR2BucketLike(bindingBucket)) {
      return { bucket: bindingBucket as R2Bucket, name: bindingName };
    }
  }

  return null;
}

function isR2BucketLike(value: unknown): value is R2Bucket {
  return Boolean(value) &&
    typeof (value as R2Bucket).put === "function" &&
    typeof (value as R2Bucket).get === "function";
}

async function loadFileTextFromBucket(bucket: R2Bucket, file: FileReference) {
  const candidates = buildTextKeyCandidates(file);
  console.log("Building OCR candidates for", {
    bucket: file.bucket,
    key: file.key,
    textKey: file.textKey,
    candidateKeys: candidates.map(candidate => candidate.key),
  });
  for (const candidate of candidates) {
    try {
      console.log("Attempting to load candidate", candidate);
      const object = await bucket.get(candidate.key);
      if (!object) {
        console.log("Candidate missing from bucket", { key: candidate.key });
        continue;
      }
      const contentType = object.httpMetadata?.contentType || "";
      const treatAsText = candidate.source === "ocr" || isTextKeyByExtension(candidate.key);
      if (!treatAsText && !isTextContentType(contentType)) {
        console.log("Skipping non-text candidate", { key: candidate.key, contentType });
        continue;
      }
      const text = await object.text();
      if (!text.trim()) {
        console.log("Candidate produced empty text", { key: candidate.key });
        continue;
      }
      console.log("OCR load success:", { key: candidate.key, textLength: text.length, source: candidate.source });
      return { text, source: candidate.source as "ocr" | "original", key: candidate.key };
    } catch (err) {
      console.warn(`Failed to load ${candidate.key} for ${file.bucket}/${file.key}`, err);
      continue;
    }
  }
  console.log("OCR load failure for all candidates", {
    bucket: file.bucket,
    key: file.key,
    textKey: file.textKey,
    candidates: candidates.map(candidate => candidate.key),
  });
  return null;
}

function buildTextKeyCandidates(file: FileReference) {
  const seen = new Set<string>();
  const candidates: Array<{ key: string; source: "ocr" | "original" }> = [];
  const addCandidate = (key: string, source: "ocr" | "original") => {
    if (seen.has(key)) return;
    seen.add(key);
    candidates.push({ key, source });
  };

  for (const key of buildOcrCandidateKeys(file)) {
    addCandidate(key, "ocr");
  }
  for (const key of buildOriginalTextCandidates(file)) {
    addCandidate(key, "original");
  }
  return candidates;
}

function buildOcrCandidateKeys(file: FileReference) {
  const trimmedKey = file.key?.trim() || "";
  const trimmedTextKey = file.textKey?.trim() || "";
  const base = stripExtensions(trimmedKey);
  const textKeyBase = stripExtensions(trimmedTextKey);
  const baseLower = base.toLowerCase();
  return uniqueStrings([
    trimmedTextKey,
    trimmedTextKey ? `${trimmedTextKey}.ocr.txt` : null,
    trimmedTextKey ? `${trimmedTextKey}.txt` : null,
    trimmedKey ? `${trimmedKey}.ocr.txt` : null,
    trimmedKey ? `${trimmedKey}.txt` : null,
    base ? `${base}.ocr.txt` : null,
    base ? `${base}.txt` : null,
    base ? `ocr/${base}.txt` : null,
    base ? `ocr/${base}.ocr.txt` : null,
    base ? `transcripts/${base}.txt` : null,
    base ? `transcripts/${base}.ocr.txt` : null,
    textKeyBase ? `${textKeyBase}.ocr.txt` : null,
    textKeyBase ? `${textKeyBase}.txt` : null,
    baseLower ? `${baseLower}.txt` : null,
    baseLower ? `${baseLower}.ocr.txt` : null,
  ]);
}

function buildOriginalTextCandidates(file: FileReference) {
  const trimmedKey = file.key?.trim() || "";
  const trimmedTextKey = file.textKey?.trim() || "";
  const base = stripExtensions(trimmedKey);
  const candidates: Array<string | null> = [
    isTextKeyByExtension(trimmedKey) ? trimmedKey : null,
    isTextKeyByExtension(trimmedTextKey) ? trimmedTextKey : null,
    base ? `${base}.txt` : null,
    trimmedKey ? `${trimmedKey}.txt` : null,
  ];
  return uniqueStrings(candidates);
}

function stripExtensions(value?: string | null) {
  if (!value) return "";
  return value
    .replace(/\.ocr\.txt$/i, "")
    .replace(/\.txt$/i, "")
    .replace(/\.pdf$/i, "");
}

function uniqueStrings(values: Array<string | null | undefined>) {
  const set = new Set<string>();
  const results: string[] = [];
  for (const value of values) {
    const trimmed = value?.trim();
    if (!trimmed) continue;
    if (set.has(trimmed)) continue;
    set.add(trimmed);
    results.push(trimmed);
  }
  return results;
}

function isTextContentType(contentType: string) {
  const lowered = contentType.toLowerCase();
  if (!lowered) return false;
  return lowered.startsWith("text/") ||
    lowered.includes("json") ||
    lowered.includes("csv") ||
    lowered.includes("xml");
}

function isTextKeyByExtension(key: string) {
  return /\.(txt|csv|json|xml|md|markdown|html|log)$/i.test(key);
}

async function attachFileToVectorStore(env: Env, storeId: string, fileId: string) {
  await openAIJsonFetch(env, `/vector_stores/${storeId}/files`, {
    method: "POST",
    json: { file_id: fileId },
  });
  return waitForVectorStoreFile(env, storeId, fileId);
}

async function waitForVectorStoreFile(env: Env, storeId: string, fileId: string) {
  const deadline = Date.now() + VECTOR_POLL_TIMEOUT_MS;
  let lastStatus = "queued";
  while (Date.now() < deadline) {
    const record = await fetchVectorStoreFileStatus(env, storeId, fileId);
    if (record.status === "completed") {
      return "completed";
    }
    if (record.status === "failed") {
      const message = record.last_error || "Vector store indexing failed.";
      throw new Error(message);
    }
    lastStatus = record.status || lastStatus;
    await delay(VECTOR_POLL_INTERVAL_MS);
  }
  return lastStatus || "processing";
}

async function fetchVectorStoreFileStatus(env: Env, storeId: string, fileId: string) {
  const data = await openAIJsonFetch(env, `/vector_stores/${storeId}/files/${fileId}`, { method: "GET" });
  return {
    status: typeof data?.status === "string" ? data.status : undefined,
    last_error: typeof data?.last_error?.message === "string" ? data.last_error.message : undefined,
  };
}

async function getPersistedVectorStoreId(env: Env) {
  if (env.VECTOR_STORE_ID?.trim()) return env.VECTOR_STORE_ID.trim();
  if (inMemoryVectorStoreId) return inMemoryVectorStoreId;
  if (env.DOCS_KV) {
    const stored = await env.DOCS_KV.get(VECTOR_STORE_CACHE_KEY);
    if (stored) {
      inMemoryVectorStoreId = stored;
      return stored;
    }
  }
  return null;
}

async function getOrCreateVectorStoreId(env: Env, label?: string) {
  const existing = await getPersistedVectorStoreId(env);
  if (existing) return existing;
  const now = new Date().toISOString();
  const name = label ? `OWEN • ${label}` : "OWEN Knowledge Base";
  const store = await openAIJsonFetch(env, "/vector_stores", {
    method: "POST",
    json: {
      name: name.slice(0, 60),
      metadata: { created_at: now },
    },
  });
  const storeId = typeof store?.id === "string" ? store.id : "";
  if (!storeId) throw new Error("OpenAI did not return a vector_store id.");
  inMemoryVectorStoreId = storeId;
  if (env.DOCS_KV) {
    await env.DOCS_KV.put(VECTOR_STORE_CACHE_KEY, storeId);
  }
  return storeId;
}

type OpenAIFetchOptions = RequestInit & { json?: unknown };

async function openAIJsonFetch(env: Env, path: string, options: OpenAIFetchOptions = {}) {
  const base = env.OPENAI_API_BASE?.replace(/\/$/, "") || "https://api.openai.com/v1";
  const url = path.startsWith("http") ? path : `${base}${path}`;
  const { json, headers, body, ...rest } = options;
  const init: RequestInit = {
    ...rest,
    headers: {
      authorization: `Bearer ${env.OPENAI_API_KEY}`,
      "content-type": "application/json",
      ...(headers || {}),
    },
    body: json !== undefined ? JSON.stringify(json) : body,
  };
  const resp = await fetch(url, init);
  const data = await safeJson(resp);
  if (!resp.ok) {
    const message = data?.error?.message || data?.message || resp.statusText || "OpenAI request failed.";
    throw new Error(message);
  }
  return data;
}

function delay(ms: number) {
  return new Promise<void>(resolve => setTimeout(resolve, ms));
}

async function retryOpenAI(fn: (attempt: number) => Promise<Response>, label: string): Promise<Response> {
  let lastResp: Response | null = null;
  for (let i = 0; i < RETRY_DELAYS_MS.length; i++) {
    if (i > 0) {
      const delayMs = RETRY_DELAYS_MS[i] ?? 0;
      await delay(delayMs);
    }
    const resp = await fn(i);
    lastResp = resp;
    if (resp.ok) return resp;
    const status = resp.status;
    const retryable = status === 429 || (status >= 500 && status < 600);
    if (!retryable || i === RETRY_DELAYS_MS.length - 1) {
      return resp;
    }
    try {
      const clone = resp.clone();
      const body = await safeJson(clone);
      console.warn("[Retry] OpenAI retry", { label, attempt: i + 1, status, body });
    } catch {
      console.warn("[Retry] OpenAI retry", { label, attempt: i + 1, status });
    }
  }
  return lastResp as Response;
}

function formatDocumentSummaries(list: DocSummary[]) {
  return list
    .map(item => {
      const title = item.filename || item.file_id || "Attachment";
      const points = Array.isArray(item.key_points) && item.key_points.length
        ? `Key Points:\n- ${item.key_points.join("\n- ")}`
        : "";
      const summary = item.summary ? `Summary: ${item.summary}` : "";
      const excerpt = item.raw_excerpt ? `Excerpt: ${item.raw_excerpt}` : "";
      return [`Document: ${title}`, summary, points, excerpt].filter(Boolean).join("\n");
    })
    .join("\n\n");
}

function extractOutputText(payload: any): string {
  if (!payload) return "";
  const stringifyContent = (content: any[]): string => {
    return content
      .map((part: any) => {
        if (typeof part?.text === "string") return part.text;
        if (part && typeof part.json === "object") return JSON.stringify(part.json);
        return "";
      })
      .filter(Boolean)
      .join("");
  };

  if (Array.isArray(payload.output)) {
    const joined = payload.output
      .map((item: any) =>
        Array.isArray(item?.content)
          ? stringifyContent(item.content)
          : "",
      )
      .filter(Boolean)
      .join("\n");
    if (joined.trim()) return joined;
  }
  if (Array.isArray(payload.output_text)) {
    const joined = payload.output_text.filter(Boolean).join("\n");
    if (joined.trim()) return joined;
  }
  if (Array.isArray(payload.response?.output)) {
    const joined = payload.response.output
      .map((item: any) =>
        Array.isArray(item?.content) ? stringifyContent(item.content) : "",
      )
      .filter(Boolean)
      .join("\n");
    if (joined.trim()) return joined;
  }
  if (typeof payload.response?.output_text === "string") {
    return payload.response.output_text;
  }
  if (Array.isArray(payload.response?.output_text)) {
    const joined = payload.response.output_text.filter(Boolean).join("\n");
    if (joined.trim()) return joined;
  }
  return "";
}

async function appendMetaTags(env: Env, tags: string[]) {
  if (!env.DOCS_KV) return;
  const normalized = normalizeTags(tags);
  if (!normalized.length) return;
  const key = "meta_tags";
  let existing: string[] = [];
  try {
    const raw = await env.DOCS_KV.get(key);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed?.tags)) {
        existing = parsed.tags.map((tag: unknown) => (typeof tag === "string" ? tag : "")).filter(Boolean);
      }
    }
  } catch {}
  const union = new Set(existing);
  normalized.forEach(tag => union.add(tag));
  const finalTags = Array.from(union).slice(-500);
  await env.DOCS_KV.put(key, JSON.stringify({ tags: finalTags, updated_at: new Date().toISOString() }));
  const timelineKey = `meta_tags:${Date.now()}`;
  await env.DOCS_KV.put(timelineKey, JSON.stringify({ tags: normalized, ts: new Date().toISOString() }), {
    expirationTtl: 60 * 60 * 24 * 30,
  });
}

function coerceLectureTags(raw: unknown) {
  const counts = new Map<string, number>();
  if (!Array.isArray(raw)) return [];
  raw.forEach((entry) => {
    if (typeof entry === "string") {
      const normalized = normalizeTags([entry]);
      const tag = normalized.length ? normalized[0] : "";
      if (tag) counts.set(tag, (counts.get(tag) || 0) + 1);
      return;
    }
    if (entry && typeof entry === "object") {
      const tagRaw = typeof (entry as any).tag === "string" ? (entry as any).tag : "";
      const normalized = normalizeTags([tagRaw]);
      const tag = normalized.length ? normalized[0] : "";
      const count = typeof (entry as any).count === "number" ? Math.max(1, Math.floor((entry as any).count)) : 1;
      if (tag) counts.set(tag, (counts.get(tag) || 0) + count);
    }
  });
  return Array.from(counts.entries()).map(([tag, count]) => ({ tag, count }));
}

async function loadLectureAnalytics(env: Env, docId: string): Promise<LectureAnalytics> {
  const fallback = { docId, tags: [], updated_at: "" };
  try {
    if (!env.OWEN_ANALYTICS) return fallback;
    const aggregate = await aggregateLectureAnalytics({
      bucket: env.OWEN_ANALYTICS,
      lectureId: docId,
      days: 365,
    });
    const questionsAgg = await loadTopQuestionsForLecture({
      bucket: env.OWEN_ANALYTICS,
      lectureId: docId,
      limit: 10,
    });
    const updates = [aggregate.lastUpdated, questionsAgg?.lastUpdated].filter(Boolean) as string[];
    const lastUpdated = updates.sort().slice(-1)[0] || "";
    return {
      docId,
      tags: aggregate.topTags,
      updated_at: lastUpdated,
      entities: aggregate.topEntities,
      questions: questionsAgg?.top ?? [],
      lastUpdated: lastUpdated || undefined,
    };
  } catch (err) {
    console.error("[ANALYTICS_READ_FAILED]", { docId, error: err instanceof Error ? err.message : String(err) });
    return fallback;
  }
}

async function appendLectureAnalytics(env: Env, docId: string, rawTags: string[]): Promise<LectureAnalytics> {
  if (!rawTags.length) {
    return { docId, tags: [], updated_at: "" };
  }
  return loadLectureAnalytics(env, docId);
}

function normalizeTags(tags: string[]) {
  return tags
    .map(tag => tag.trim())
    .map(tag => (tag.startsWith("#") ? tag : `#${tag}`))
    .map(tag => tag.replace(/[^#a-z0-9_-]/gi, "").toLowerCase())
    .filter(Boolean);
}

function buildOcrKey(file: FileReference) {
  const base = stripExtensions(file.textKey?.trim() || file.key);
  return base ? `${base}.ocr.txt` : `${file.key}.ocr.txt`;
}

async function mirrorBucketObjectToOpenAI(bucket: R2Bucket, key: string, env: Env) {
  const object = await bucket.get(key);
  if (!object || !object.body) {
    throw new Error("Object not found in bucket for OCR mirroring.");
  }
  const contentType = object.httpMetadata?.contentType || "application/octet-stream";
  const filename = sanitizeFilename(key.split("/").pop() || "upload.bin");
  const buffer = await object.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  // Assistants upload for text/OCR
  const assistantsFile = new File([buffer], filename, { type: contentType });
  const form = new FormData();
  form.append("purpose", "assistants");
  form.append("file", assistantsFile, filename);
  const resp = await fetch(`${env.OPENAI_API_BASE}/files`, {
    method: "POST",
    headers: { authorization: `Bearer ${env.OPENAI_API_KEY}` },
    body: form,
  });
  const data = await safeJson(resp);
  if (!resp.ok) {
    const message = data?.error?.message || "OpenAI Files upload failed for OCR fallback.";
    throw new Error(message);
  }
  const uploadedId = typeof data?.id === "string" ? data.id : "";
  if (!uploadedId) {
    throw new Error("OpenAI did not return a file id during OCR fallback upload.");
  }

  // Vision upload (best effort for images)
  let visionFileId: string | undefined;
  const wantsVision = contentType.toLowerCase().startsWith("image/") || isLikelyImageFilename(filename);
  const imageSignature = detectImageSignature(bytes);
  const imageHead = bytesToBase64(bytes.subarray(0, Math.min(48, bytes.length))).slice(0, 12);
  if (wantsVision) {
    if (imageSignature === "unknown") {
      console.warn("Vision mirror upload skipped: invalid image signature", {
        key,
        head: imageHead,
        mimeType: contentType,
      });
    } else {
      try {
        const visionMime = normalizeOcrImageMime(contentType, imageSignature);
        logOcrInputDebug({
          key,
          head: imageHead,
          mimeType: visionMime,
          signature: imageSignature,
          byteLength: bytes.length,
        });
        const visionFile = new File([buffer], filename, { type: visionMime });
        const visionForm = new FormData();
        visionForm.append("purpose", "vision");
        visionForm.append("file", visionFile, filename);
        const visionResp = await fetch(`${env.OPENAI_API_BASE}/files`, {
          method: "POST",
          headers: { authorization: `Bearer ${env.OPENAI_API_KEY}` },
          body: visionForm,
        });
        const visionData = await safeJson(visionResp);
        if (visionResp.ok && typeof visionData?.id === "string") {
          visionFileId = visionData.id;
        } else {
          const msg = visionData?.error?.message || visionData?.message || visionResp.statusText || "Vision upload failed.";
          console.warn("Vision mirror upload failed (non-fatal).", { key, msg });
        }
      } catch (err) {
        console.warn("Vision mirror upload threw (non-fatal).", { key, err: String(err) });
      }
    }
  }

  return { fileId: uploadedId, visionFileId, filename };
}

async function storeConversationTags(env: Env, conversationId: string, tags: string[]) {
  if (!env.DOCS_KV) return;
  const normalized = normalizeTags(tags);
  if (!normalized.length) return;
  const kvKey = `conv:${conversationId}:tags`;
  const payload = {
    conversation_id: conversationId,
    tags: normalized,
    updated_at: new Date().toISOString(),
  };
  await env.DOCS_KV.put(kvKey, JSON.stringify(payload));
}

async function handleMetaTagsRequest(req: Request, env: Env): Promise<Response> {
  if (!env.DOCS_KV) {
    return json({ tags: [] });
  }
  const url = new URL(req.url);
  const conversationId = url.searchParams.get("conversation_id");
  if (!conversationId) {
    return json({ error: "conversation_id required" }, 400);
  }
  const record = await loadConversationTags(env, conversationId);
  return json(record);
}
function debugFormatLog(message: string, payload?: Record<string, unknown>) {
  if (!DEBUG_FORMATTING) return;
  console.log(`[FORMAT] ${message}`, payload || {});
}

function finalizeAssistantText(rawText: string, formattedText: string, fallbackPreview?: string): string {
  const formatted = (formattedText || "").trim();
  if (formatted) return formatted;
  const raw = (rawText || "").trim();
  if (raw) return raw;
  if (fallbackPreview) return `(empty assistant message — showing preview)\n\n${fallbackPreview}`;
  return "(empty assistant message — formatting fallback engaged)";
}
