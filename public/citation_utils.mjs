/**
 * Client-side citation utilities for rendering source pills/links.
 *
 * Used by: `public/chat.js` and citation-related UI to normalize and render citations.
 *
 * Key exports:
 * - Domain and URL validation helpers.
 * - Map builders for citation segments and source lists.
 * - DOM helpers for rendering and resolving citation pills.
 *
 * Assumptions:
 * - Citation segments follow the Responses API shape and include numeric ids + URLs.
 * - In-app attachment links may be relative (leading slash).
 */
/**
 * Extract a hostname/domain from a URL string.
 *
 * @param url - URL string to parse.
 * @returns Hostname or a best-effort fallback.
 */
export function getCitationDomain(url) {
  try {
    return new URL(url).hostname || url;
  } catch {
    return url.replace(/^https?:\/\//i, "").split("/")[0] || url;
  }
}

/**
 * Heuristic check for URLs that are safe to show as citations.
 *
 * @param url - URL string to validate.
 * @returns True when the URL passes blocklist and format checks.
 */
export function isLikelyValidCitationUrl(url) {
  if (!url) return false;
  if (url.startsWith("/#")) return false;
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

/**
 * Build a citation id -> metadata map from answer segments.
 *
 * @param segments - Answer segments that include citation entries.
 * @returns Map keyed by citation id.
 */
export function buildCitationMapFromSegments(segments) {
  const map = new Map();
  (segments || []).forEach(seg => {
    if (!seg || seg.type !== "citation") return;
    const id = Number(seg.id);
    if (!Number.isFinite(id)) return;
    const url = typeof seg.url === "string" ? seg.url : "";
    if (!url) return;
    if (!map.has(id)) {
      map.set(id, { id, url, title: typeof seg.title === "string" ? seg.title : "" });
    }
  });
  return map;
}

/**
 * Normalize a list of cited sources into a consistent shape.
 *
 * @param sources - Source list from the backend (optional).
 * @param citationMap - Map of citation ids to URL metadata.
 * @returns Sorted list of normalized source objects.
 */
export function normalizeCitedSources(sources, citationMap) {
  const normalized = [];
  if (Array.isArray(sources) && sources.length) {
    sources.forEach(src => {
      const id = Number(src?.id);
      const url = typeof src?.url === "string" ? src.url : "";
      if (!Number.isFinite(id) || !url) return;
      normalized.push({
        id,
        url,
        title: typeof src?.title === "string" ? src.title : "",
        domain: typeof src?.domain === "string" && src.domain ? src.domain : getCitationDomain(url),
        snippet: typeof src?.snippet === "string" ? src.snippet : (typeof src?.excerpt === "string" ? src.excerpt : ""),
      });
    });
  } else if (citationMap instanceof Map) {
    citationMap.forEach((value, key) => {
      let url = "";
      let id = Number.NaN;
      let title = "";
      let snippet = "";
      let domain = "";
      if (typeof value === "string") {
        url = value;
        id = Number(key);
      } else if (value && typeof value === "object") {
        url = typeof value.url === "string" ? value.url : "";
        title = typeof value.title === "string" ? value.title : "";
        snippet = typeof value.snippet === "string" ? value.snippet : "";
        domain = typeof value.domain === "string" ? value.domain : "";
        const valueId = value.id ?? key;
        id = Number(valueId);
      }
      if (!Number.isFinite(id) || !url) return;
      normalized.push({
        id,
        url,
        title,
        domain: domain || getCitationDomain(url),
        snippet,
      });
    });
  }
  normalized.sort((a, b) => a.id - b.id);
  return normalized;
}

/**
 * Build a citation pill DOM element for inline rendering.
 *
 * @param cite - Citation metadata (id, url, title/domain).
 * @returns A `<sup>` element containing a link or disabled button.
 * @remarks Side effects: creates DOM nodes.
 */
export function buildCitationPill(cite) {
  const id = Number(cite?.id);
  const url = typeof cite?.url === "string" ? cite.url : "";
  if (!Number.isFinite(id)) return document.createTextNode("");
  const label = cite?.title || cite?.domain || url || "Source unavailable";
  const wrap = document.createElement("sup");
  wrap.className = "citation-sup";
  wrap.dataset.citeInline = "true";
  if (!url || !isLikelyValidCitationUrl(url)) {
    const pill = document.createElement("button");
    pill.type = "button";
    pill.className = "citation-pill citation-chip citation-pill--disabled";
    pill.textContent = `[${id}]`;
    pill.dataset.citeId = String(id);
    pill.dataset.citeInline = "true";
    pill.setAttribute("aria-label", `Open source ${id}: ${label}`);
    pill.setAttribute("aria-disabled", "true");
    pill.title = "Source unavailable";
    pill.disabled = true;
    wrap.appendChild(pill);
    return wrap;
  }
  const pill = document.createElement("a");
  pill.className = "citation-pill citation-chip";
  pill.href = url;
  pill.target = "_blank";
  pill.rel = "noopener noreferrer";
  pill.textContent = `[${id}]`;
  pill.dataset.citeId = String(id);
  pill.dataset.citeInline = "true";
  pill.setAttribute("aria-label", `Open source ${id}: ${label}`);
  pill.title = label;
  wrap.appendChild(pill);
  return wrap;
}

/**
 * Replace citation placeholders in text nodes with pill elements.
 *
 * @param root - Root DOM node to scan.
 * @param citationMap - Map of citation ids to metadata.
 * @param placeholderPrefix - Prefix used in placeholder markers.
 * @remarks Side effects: mutates the DOM under `root`.
 */
export function replaceCitationPlaceholders(root, citationMap, placeholderPrefix) {
  if (!root || !(citationMap instanceof Map) || !citationMap.size) return;
  const prefix = placeholderPrefix || "";
  const pattern = new RegExp(`${prefix}(\\d+)__`, "g");
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) {
    const node = walker.currentNode;
    if (node?.nodeValue && prefix && node.nodeValue.includes(prefix)) {
      nodes.push(node);
    }
  }
  nodes.forEach(node => {
    const text = node.nodeValue || "";
    let lastIndex = 0;
    const frag = document.createDocumentFragment();
    for (const match of text.matchAll(pattern)) {
      const index = match.index ?? 0;
      if (index > lastIndex) {
        frag.appendChild(document.createTextNode(text.slice(lastIndex, index)));
      }
      const id = Number(match[1]);
      const cite = citationMap.get(id) ?? citationMap.get(String(id));
      frag.appendChild(cite ? buildCitationPill(cite) : document.createTextNode(match[0]));
      lastIndex = index + match[0].length;
    }
    if (lastIndex < text.length) {
      frag.appendChild(document.createTextNode(text.slice(lastIndex)));
    }
    if (node.parentNode) {
      node.parentNode.replaceChild(frag, node);
    }
  });
}

function disableCitationLink(link, label) {
  link.removeAttribute("href");
  link.removeAttribute("target");
  link.removeAttribute("rel");
  link.classList.add("cite--disabled");
  link.setAttribute("aria-disabled", "true");
  link.title = label || "Source unavailable";
}

/**
 * Resolve and validate citation links in the DOM, disabling invalid ones.
 *
 * @param root - Root DOM node to scan.
 * @param citationMap - Map of citation ids to metadata.
 * @remarks Side effects: mutates anchor elements under `root`.
 */
export function resolveCitationLinks(root, citationMap) {
  if (!root) return;
  root.querySelectorAll("a.cite[data-cite-id]").forEach(link => {
    const id = link.getAttribute("data-cite-id");
    const value = id && citationMap instanceof Map
      ? citationMap.get(id) ?? citationMap.get(Number(id))
      : null;
    const url = typeof value === "string" ? value : value?.url;
    const label = value?.title || value?.domain || url || "Source unavailable";
    if (id) {
      link.dataset.citeId = String(id);
      link.dataset.citeInline = "true";
    }
    if (url && isLikelyValidCitationUrl(url)) {
      link.href = url;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.setAttribute("aria-disabled", "false");
      link.setAttribute("aria-label", `Open source ${id}: ${label}`);
      link.classList.remove("cite--disabled");
      link.title = label;
      return;
    }
    disableCitationLink(link);
    if (id) {
      link.setAttribute("aria-label", `Open source ${id}: ${label}`);
    }
  });
}
