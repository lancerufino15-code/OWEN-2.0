/**
 * Cloudflare Worker environment bindings for OWEN.
 *
 * Used by: `src/index.ts` and analytics/diagnostics helpers to access secrets, KV,
 * and R2 buckets.
 *
 * Key exports:
 * - `Env`: binding contract for Worker runtime configuration.
 *
 * Assumptions:
 * - Secrets are provided via Wrangler (`wrangler secret put`).
 * - Optional flags are passed as strings ("1", "true", etc.) and parsed by callers.
 */
export interface Env {
  /** Static asset binding for the worker (served from /public). */
  ASSETS: Fetcher;
  /** Workers AI binding for schema-constrained model calls. */
  AI?: Ai;
  /** OpenAI API key used for Responses/Chat calls. */
  OPENAI_API_KEY: string;
  /** Base URL for OpenAI-compatible REST API. */
  OPENAI_API_BASE: string;
  /** Optional override for quiz AI timeout (milliseconds). */
  QUIZ_AI_TIMEOUT_MS?: string;
  /** Optional override for quiz Workers AI model id. */
  QUIZ_AI_MODEL?: string;
  /** Legacy or alternate quiz model override. */
  QUIZ_MODEL?: string;
  /** Optional default model override when the UI does not specify one. */
  DEFAULT_MODEL?: string;
  /** If set, strips sampling params for stricter deterministic payloads. */
  OWEN_STRIP_SAMPLING_PARAMS?: string;
  /** Enables extra diagnostics and debug logging when truthy. */
  OWEN_DEBUG?: string;
  /** When set, preserves raw study guide JSON for inspection. */
  DEBUG_STUDY_GUIDE_JSON?: string;
  /** Passcode used to unlock faculty-only UI and endpoints. */
  FACULTY_PASSCODE?: string;
  /** Frontend passcode variant for Next.js style builds (if used). */
  NEXT_PUBLIC_FACULTY_PASSCODE?: string;
  /** Frontend passcode variant for Vite builds (if used). */
  VITE_FACULTY_PASSCODE?: string;
  /** Soft target for minimum unique sources in free-response mode. */
  FREE_RESPONSE_MIN_UNIQUE_SOURCES?: string;
  /** Alias for minimum distinct sources (legacy or alternate config). */
  MIN_DISTINCT_SOURCES?: string;
  /** If truthy, hard-enforces minimum distinct sources. */
  ENFORCE_MIN_DISTINCT_SOURCES?: string;
  /** Free-response specific enforcement toggle. */
  FREE_RESPONSE_ENFORCE_MIN_UNIQUE_SOURCES?: string;
  /** Optional persistent vector store id for retrieval calls. */
  VECTOR_STORE_ID?: string;
  /** KV namespace for doc metadata and cached artifacts. */
  DOCS_KV?: KVNamespace;
  /** KV namespace for diagnostics (request snapshots, analytics key, etc.). */
  OWEN_DIAG_KV?: KVNamespace;
  /** Model id override for GPT-4 (chat). */
  GPT4_MODEL_ID?: string;
  /** Model id override for GPT-4 Turbo (chat). */
  GPT4_TURBO_MODEL_ID?: string;
  /** Model id override for GPT-5 (responses). */
  GPT5_MODEL_ID?: string;
  /** Model id override for image generation (gpt-image-1). */
  GPT_IMAGE_1_MODEL_ID?: string;
  /** Model id override for mini image generation. */
  GPT_IMAGE_1_MINI_MODEL_ID?: string;
  /** Model id override for DALL-E 3. */
  DALLE3_MODEL_ID?: string;
  /** Model id override for DALL-E 2. */
  DALLE2_MODEL_ID?: string;
  /** Model id override for OCR pipeline. */
  OCR_MODEL_ID?: string;
  /** R2 bucket for primary document storage. */
  OWEN_BUCKET: R2Bucket;
  /** R2 bucket for ingest staging. */
  OWEN_INGEST: R2Bucket;
  /** R2 bucket for notes or curated content. */
  OWEN_NOTES: R2Bucket;
  /** R2 bucket for test fixtures. */
  OWEN_TEST: R2Bucket;
  /** R2 bucket for user uploads (default). */
  OWEN_UPLOADS: R2Bucket;
  /** R2 bucket for analytics exports. */
  OWEN_ANALYTICS: R2Bucket;
  /** Legacy bucket binding (typo or alternate name). */
  OWN_INGEST: R2Bucket;
}
