/**
 * Text hygiene utilities for OCR and transcript processing.
 *
 * Used by: `src/index.ts` ingestion flow and tests to clean slide text before
 * study-guide prompts or storage.
 *
 * Key exports:
 * - `REFUSAL_PATTERNS`: Regex list used to drop model refusal lines.
 * - `sanitizeSlideText` / `sanitizeDocText`: Normalizers for slide-style text.
 *
 * Assumptions:
 * - Input is raw OCR or transcript text with possible "Slide X" headers.
 * - Returns placeholder "[NO TEXT]" for empty slide bodies.
 */
export const REFUSAL_PATTERNS: RegExp[] = [
  /i\s+can't\s+assist/i,
  /i\s+can\'t\s+assist/i,
  /i\s+cannot\s+assist/i,
  /i\s*'m\s+sorry/i,
  /unable\s+to/i,
  /cannot\s+comply/i,
  /as\s+an\s+ai/i,
  /i\s+cannot\s+help\s+with\s+that/i,
];

function normalizeApostrophes(text: string): string {
  return text.replace(/[\u2019\u2018\u201b`]/g, "'");
}

function normalizeLineEndings(text: string): string {
  return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
}

/**
 * Normalize slide text: fix apostrophes, trim whitespace, remove refusals.
 *
 * @param raw - Raw slide or transcript text.
 * @returns Cleaned text with consecutive blank lines collapsed.
 */
export function sanitizeSlideText(raw: string): string {
  const normalized = normalizeApostrophes(normalizeLineEndings(raw || ""));
  const lines = normalized.split("\n");
  const filtered = lines.filter((line) => {
    return !REFUSAL_PATTERNS.some((re) => re.test(line));
  });
  const trimmed = filtered.map((line) => line.replace(/[ \t]+$/g, ""));
  const collapsed: string[] = [];
  let blankStreak = 0;
  for (const line of trimmed) {
    if (line.trim() === "") {
      blankStreak += 1;
      if (blankStreak > 1) continue;
      collapsed.push("");
      continue;
    }
    blankStreak = 0;
    collapsed.push(line);
  }
  return collapsed.join("\n").trim();
}

/**
 * Normalize a full document with optional "Slide N" boundaries.
 *
 * @param raw - Raw document text (possibly OCR output).
 * @returns Cleaned text with slide blocks separated by blank lines.
 */
export function sanitizeDocText(raw: string): string {
  const normalized = normalizeLineEndings(raw || "");
  const slideRegex = /^Slide\s+\d+\s+\(p\.\d+\):.*$/gim;
  const matches = Array.from(normalized.matchAll(slideRegex));

  if (!matches.length) {
    const cleaned = sanitizeSlideText(normalized);
    return cleaned || "[NO TEXT]";
  }

  const blocks: string[] = [];
  const firstMatch = matches[0];
  if (firstMatch && typeof firstMatch.index === "number" && firstMatch.index > 0) {
    const prefix = normalized.slice(0, firstMatch.index);
    const cleanedPrefix = sanitizeSlideText(prefix);
    if (cleanedPrefix) {
      blocks.push(cleanedPrefix);
    }
  }

  for (let i = 0; i < matches.length; i += 1) {
    const match = matches[i];
    if (!match || typeof match.index !== "number") continue;
    const header = match[0].trimEnd();
    const headerEnd = match.index + match[0].length;
    const nextIndex = i + 1 < matches.length && typeof matches[i + 1].index === "number"
      ? (matches[i + 1].index as number)
      : normalized.length;
    let bodyRaw = normalized.slice(headerEnd, nextIndex);
    if (bodyRaw.startsWith("\n")) bodyRaw = bodyRaw.slice(1);
    const cleanedBody = sanitizeSlideText(bodyRaw);
    const finalBody = cleanedBody || "[NO TEXT]";
    blocks.push(`${header}\n${finalBody}`.trimEnd());
  }

  return blocks.join("\n\n");
}
