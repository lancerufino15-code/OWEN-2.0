/**
 * Purpose:
 * - Utilities for the documents library index persisted in R2.
 *
 * Responsibilities:
 * - Compute deterministic docIds, manifest paths, and cache keys for library ingestion.
 * - Normalize titles/tokens and score records for /api/library/* search endpoints.
 * - Persist/read JSONL index files within the Workers R2 bindings.
 *
 * Architecture role:
 * - Shared helper module used by the worker entrypoint to manage the lecture library cache.
 *
 * Constraints:
 * - Runs in Cloudflare Workers (Web Crypto + R2 only; no Node buffers or filesystem access).
 */
import { normalizePlainText } from "./pdf/normalize";

/** Metadata stored per document in the library index JSONL file. */
export type LibraryIndexRecord = {
  docId: string;
  bucket: string;
  key: string;
  title: string;
  normalizedTokens: string[];
  hashBasis: string;
  hashFieldsUsed?: string[];
  etag?: string | null;
  size?: number;
  uploaded?: string;
  status?: "ready" | "missing" | "needs_browser_ocr";
  preview?: string;
  manifestKey?: string;
  extractedKey?: string;
};

export const LIBRARY_INDEX_KEY = "library/index.jsonl";
export const LIBRARY_INDEX_PREFIX = "index/";
export const LIBRARY_QUEUE_PREFIX = "queue/scanned/";
export const EXTRACTED_PREFIX = "extracted/";
export const MANIFEST_PREFIX = "manifests/";
export const LIBRARY_TXT_PREFIX = "library_txt/";
export const LIBRARY_IMAGES_PREFIX = "library/images/";
export const LIBRARY_CONVERTED_PREFIX = "converted-images/";

const enc = new TextEncoder();

/** Convert an ArrayBuffer digest to a lowercase hex string. */
function toHex(buffer: ArrayBuffer) {
  return Array.from(new Uint8Array(buffer))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

/** Compute a SHA-256 digest over the provided value using Web Crypto (Workers-safe). */
export async function sha256(value: string) {
  const digest = await crypto.subtle.digest("SHA-256", enc.encode(value));
  return toHex(digest);
}

/**
 * Build a deterministic hash basis for a document using R2 metadata.
 * Prefers etag, otherwise falls back to size + uploaded timestamps so docIds change when content changes.
 */
export function computeHashBasis(
  bucket: string,
  key: string,
  meta: { etag?: string | null; size?: number; uploaded?: string | Date | null; lastModified?: string | Date | null },
) {
  const uploaded =
    meta.uploaded instanceof Date
      ? meta.uploaded.toISOString()
      : typeof meta.uploaded === "string"
        ? meta.uploaded
        : meta.lastModified instanceof Date
          ? meta.lastModified.toISOString()
          : typeof meta.lastModified === "string"
            ? meta.lastModified
            : "";
  const basis = [bucket, key, meta.etag || "", meta.size ?? "", uploaded].join(":");
  const fieldsUsed = meta.etag ? ["etag"] : ["size", "uploaded"];
  return { basis, fieldsUsed, uploaded };
}

/** Derive a library docId and audit info from bucket/key metadata. */
export async function computeDocId(
  bucket: string,
  key: string,
  meta: { etag?: string | null; size?: number; uploaded?: string | Date | null; lastModified?: string | Date | null },
) {
  const { basis, fieldsUsed, uploaded } = computeHashBasis(bucket, key, meta);
  const docId = await sha256(basis);
  return { docId, basis, fieldsUsed, uploaded };
}

/** Convert an R2 object key into a human-readable document title. */
export function titleFromKey(key: string) {
  const leaf = key.split("/").pop() || key;
  const withoutExt = leaf.replace(/\.[^.]+$/, "");
  const normalized = withoutExt.replace(/[_\-]+/g, " ").trim();
  return normalized || leaf || key;
}

/** Tokenize a document title into normalized search tokens for matching. */
export function tokensFromTitle(title: string) {
  const tokens = (title || "")
    .toLowerCase()
    .split(/[^a-z0-9]+/g)
    .map(tok => tok.trim())
    .filter(tok => tok.length > 1);
  return Array.from(new Set(tokens));
}

/** Detect whether the object key likely represents a PDF for ingestion/preview rules. */
export function isPdfKey(key: string) {
  return /\.pdf$/i.test(key || "");
}

/** Build the JSON index key for an individual document entry. */
export function buildIndexKeyForDoc(docId: string) {
  const safe = (docId || "").replace(/[^a-z0-9]/gi, "");
  return `${LIBRARY_INDEX_PREFIX}${safe}.json`;
}

/** Build the text extraction path for a document in the library cache. */
export function buildExtractedPath(docId: string) {
  const safe = (docId || "").replace(/[^a-z0-9]/gi, "");
  return `${EXTRACTED_PREFIX}${safe}.txt`;
}

/** Build the manifest path for a document in the library cache. */
export function buildManifestPath(docId: string) {
  const safe = (docId || "").replace(/[^a-z0-9]/gi, "");
  return `${MANIFEST_PREFIX}${safe}.json`;
}

/** Build the deterministic Machine TXT export path for a sanitized filename. */
export function buildLibraryTxtPath(filename: string) {
  const safe = (filename || "").replace(/^\/+/, "");
  return `${LIBRARY_TXT_PREFIX}${safe}`;
}

/** Build the deterministic Machine image prefix for a lecture slug. */
export function buildLibraryImagesPrefix(sanitizedTitle: string) {
  const safe = (sanitizedTitle || "").replace(/^\/+/, "").replace(/\/+$/, "");
  return `${LIBRARY_IMAGES_PREFIX}${safe}/`;
}

/** Build the deterministic image conversion prefix for a docId + format + variant descriptor. */
export function buildConvertedImagesPrefix(docId: string, format: string, variant: string) {
  const safeDoc = (docId || "").replace(/[^a-z0-9]/gi, "");
  const safeFormat = (format || "").toLowerCase() === "jpg" ? "jpg" : "png";
  const safeVariant = (variant || "").replace(/[^a-z0-9_-]/gi, "");
  const variantSuffix = safeVariant ? `${safeVariant}/` : "";
  return `${LIBRARY_CONVERTED_PREFIX}${safeDoc}/${safeFormat}/${variantSuffix}`;
}

/**
 * Read and parse the library index JSONL file from R2.
 * Returns only valid records and tolerates malformed lines to avoid breaking ingestion.
 */
export async function readIndex(bucket: R2Bucket): Promise<LibraryIndexRecord[]> {
  try {
    const object = await bucket.get(LIBRARY_INDEX_KEY);
    if (!object || !object.body) return [];
    const text = await object.text();
    return text
      .split("\n")
      .map(line => line.trim())
      .filter(Boolean)
      .map(line => {
        try {
          const parsed = JSON.parse(line);
          if (parsed && typeof parsed.docId === "string") {
            parsed.normalizedTokens = Array.isArray(parsed.normalizedTokens)
              ? parsed.normalizedTokens
              : tokensFromTitle(parsed.title || "");
            return parsed as LibraryIndexRecord;
          }
        } catch {
          // ignore malformed lines
        }
        return null;
      })
      .filter((r): r is LibraryIndexRecord => Boolean(r));
  } catch {
    return [];
  }
}

/** Persist the library index JSONL file back to R2. */
export async function writeIndex(bucket: R2Bucket, records: LibraryIndexRecord[]) {
  const lines = records
    .map(rec => JSON.stringify(rec))
    .join("\n");
  await bucket.put(LIBRARY_INDEX_KEY, lines, {
    httpMetadata: { contentType: "application/json; charset=utf-8" },
  });
}

/**
 * Score library records against a free-text query using token overlap heuristics.
 * Keeps the top matches to avoid sending excessive context to downstream chat handlers.
 */
export function scoreRecords(query: string, records: LibraryIndexRecord[], limit = 10) {
  const tokens = tokensFromTitle(query);
  const qLower = (query || "").toLowerCase();
  const scored = records.map(rec => {
    let score = 0;
    const tokenSet = new Set(rec.normalizedTokens || []);
    tokens.forEach(tok => {
      if (tokenSet.has(tok)) score += 3;
      else if (rec.normalizedTokens?.some(rt => rt.startsWith(tok))) score += 2;
      else if ((rec.title || "").toLowerCase().includes(tok)) score += 1;
    });
    if ((rec.title || "").toLowerCase().includes(qLower)) {
      score += 1;
    }
    return { rec, score };
  });
  return scored
    .sort((a, b) => b.score - a.score)
    .slice(0, limit)
    .map(entry => entry.rec);
}

/** Create a compact preview string suitable for search result snippets. */
export function normalizePreview(text?: string) {
  if (!text) return "";
  return normalizePlainText(text).slice(0, 280);
}
