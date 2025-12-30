/**
 * Citation filtering helpers for OWEN contract responses (shared module).
 *
 * Used by: front-end or shared code paths to enforce allowed source URLs.
 *
 * Key exports:
 * - `CitationGuardResult`: filtered contract plus metrics.
 * - `enforceAllowedSources`: remove claims/links outside allowed URL set.
 * - `normalizeUrl`: URL normalization helper for comparisons.
 *
 * Assumptions:
 * - Contract body/claims follow the OWEN contract schema.
 * - URLs are expected to be absolute; invalid URLs are treated as disallowed.
 */
import type { OwenContractT, OwenSourceT } from "./contract";

/**
 * Result of citation filtering, including metrics.
 */
export type CitationGuardResult = {
  contract: OwenContractT;
  metrics: {
    droppedClaims: number;
    strippedLinks: number;
  };
};

/**
 * Filter a contract to only include sources present in the allowed URL list.
 *
 * @param contract - Original OWEN contract payload.
 * @param allowedUrls - Allowed URL list (normalized internally).
 * @returns Filtered contract plus metrics about removals.
 */
export function enforceAllowedSources(contract: OwenContractT, allowedUrls: string[]): CitationGuardResult {
  const normalizedAllowed = new Set(
    (allowedUrls || [])
      .map(normalizeUrl)
      .filter((value): value is string => Boolean(value)),
  );

  if (!normalizedAllowed.size) {
    return {
      contract,
      metrics: { droppedClaims: 0, strippedLinks: 0 },
    };
  }

  const keepSource = (source: OwenSourceT) => {
    const normalized = normalizeUrl(source?.url);
    return normalized ? normalizedAllowed.has(normalized) : false;
  };

  let droppedClaims = 0;
  const filteredClaims: OwenContractT["claims"] = [];

  for (const claim of contract.claims || []) {
    const sources = (claim.sources || []).filter(keepSource);
    if (!sources.length) {
      droppedClaims += 1;
      continue;
    }
    filteredClaims.push({ ...claim, sources });
  }

  const bodyResult = stripUnapprovedLinks(contract.body_markdown || "", normalizedAllowed);
  const enrichedErrors = Array.isArray(contract.errors) ? [...contract.errors] : [];
  if (droppedClaims) {
    enrichedErrors.push(
      `Removed ${droppedClaims} unsupported citation${droppedClaims === 1 ? "" : "s"} outside the retrieved context.`,
    );
  }
  if (bodyResult.removed) {
    enrichedErrors.push(
      `Stripped ${bodyResult.removed} inline link${bodyResult.removed === 1 ? "" : "s"} lacking retrieved evidence.`,
    );
  }

  return {
    contract: {
      ...contract,
      body_markdown: bodyResult.body,
      claims: filteredClaims,
      errors: enrichedErrors,
    },
    metrics: {
      droppedClaims,
      strippedLinks: bodyResult.removed,
    },
  };
}

/**
 * Normalize a URL by stripping query/hash and normalizing path.
 *
 * @param value - URL string to normalize.
 * @returns Normalized URL string or null when invalid.
 */
export function normalizeUrl(value?: string | null): string | null {
  if (!value || typeof value !== "string") {
    return null;
  }
  try {
    const parsed = new URL(value);
    parsed.hash = "";
    parsed.search = "";
    parsed.pathname = normalizePath(parsed.pathname);
    return parsed.toString();
  } catch {
    return null;
  }
}

function normalizePath(pathname: string): string {
  if (!pathname || pathname === "/") {
    return "/";
  }
  const trimmed = pathname.replace(/\/+$/, "");
  return trimmed || "/";
}

function stripUnapprovedLinks(body: string, allowed: Set<string>) {
  if (!body) {
    return { body, removed: 0 };
  }
  let removed = 0;
  const markdownLinkRegex = /\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g;
  const cleaned = body.replace(markdownLinkRegex, (full, label, url) => {
    const normalized = normalizeUrl(url);
    if (normalized && allowed.has(normalized)) {
      return full;
    }
    removed += 1;
    return label;
  });
  return { body: cleaned, removed };
}
