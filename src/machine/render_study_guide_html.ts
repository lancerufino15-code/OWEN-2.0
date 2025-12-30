/**
 * HTML renderer for study guide outputs.
 *
 * Used by: `src/index.ts` to turn Step A/B/C outputs into a standalone HTML
 * report for debugging and exports.
 *
 * Key exports:
 * - `StepAOutput`, `StepBOutput`, `StepCOutput`: JSON shapes produced by the
 *   multi-step study guide pipeline.
 * - `renderStudyGuideHtml`: Pure renderer that returns a full HTML document.
 *
 * Assumptions:
 * - Input has already been validated by the study guide validators.
 * - All user content is escaped before insertion into HTML.
 */
import type { HighlightCategory } from "../study_guides/contracts";
import { BASE_STUDY_GUIDE_CSS, renderLegend } from "../study_guides/contracts";

type StepANumber = { value: string; unit: string };

type StepAFact = {
  text: string;
  tags: string[];
  numbers: StepANumber[];
};

type StepASection = {
  heading: string;
  facts: StepAFact[];
};

type StepATable = {
  caption: string;
  headers: string[];
  rows: string[][];
};

type StepASlide = {
  n: number;
  page: number;
  sections: StepASection[];
  tables: StepATable[];
};

type StepABuckets = {
  dx: string[];
  pathophys: string[];
  clinical: string[];
  labs: string[];
  imaging: string[];
  treatment: string[];
  complications: string[];
  risk_factors: string[];
  epidemiology: string[];
  red_flags: string[];
  buzzwords: string[];
};

type StepADiscriminator = {
  topic: string;
  signals: string[];
  pitfalls: string[];
};

type StepASourceSpan = {
  text: string;
  slides?: number[];
  pages?: number[];
};

/**
 * Step A extraction output: slide-level facts, buckets, and metadata.
 */
export type StepAOutput = {
  lecture_title: string;
  slides: StepASlide[];
  raw_facts: string[];
  buckets: StepABuckets;
  discriminators: StepADiscriminator[];
  exam_atoms: string[];
  abbrev_map: Record<string, string>;
  source_spans: StepASourceSpan[];
};

type StepBRapid = {
  clue: string;
  think_of: string;
  why: string;
  confirm: string;
};

type StepBCompareRow = {
  dx1: string;
  dx2: string;
  how_to_tell: string;
};

type StepBCompare = {
  topic: string;
  rows: StepBCompareRow[];
};

type StepBQuant = {
  item: string;
  value: string;
  note: string;
};

type StepBGlossary = {
  term: string;
  definition: string;
};

/**
 * Step B synthesis output: summaries, differentials, and glossaries.
 */
export type StepBOutput = {
  high_yield_summary: string[];
  rapid_approach_table: StepBRapid[];
  one_page_last_minute_review: string[];
  compare_differential: StepBCompare[];
  quant_cutoffs: StepBQuant[];
  pitfalls: string[];
  glossary: StepBGlossary[];
  supplemental_glue?: string[];
};

type StepCItem = { slide: number; note: string };

type HighlightClass = HighlightCategory;

type HighlightLexiconEntry = {
  phrase: string;
  normalized: string;
  className: HighlightClass;
};

/**
 * Step C QA output: omissions/conflicts plus coverage checks.
 */
export type StepCOutput = {
  unparsed_items: StepCItem[];
  omissions: StepCItem[];
  conflicts: StepCItem[];
  coverage_confidence: "Low" | "Med" | "High";
  checks: {
    has_high_yield_summary: boolean;
    has_rapid_approach_table: boolean;
    has_one_page_review: boolean;
    slide_count_stepA: number;
    slide_count_rendered: number;
  };
};

function escapeHtml(value: string): string {
  return (value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizeList<T>(value: T[] | undefined | null): T[] {
  return Array.isArray(value) ? value : [];
}

function renderList(items: string[], emptyLabel: string, renderItem?: (item: string) => string): string[] {
  if (!items.length) {
    return [`<p class="muted">${escapeHtml(emptyLabel)}</p>`];
  }
  const render = renderItem ?? ((item: string) => escapeHtml(item));
  const lines = ["<ul>"];
  for (const item of items) {
    lines.push(`  <li>${render(item)}</li>`);
  }
  lines.push("</ul>");
  return lines;
}

function renderTable(
  headers: string[],
  rows: string[][],
  tableClass = "",
  renderCell?: (cell: string) => string,
): string[] {
  const safeHeaders = normalizeList(headers).map(escapeHtml);
  const render = renderCell ?? ((cell: string) => escapeHtml(cell || ""));
  const lines = [`<table${tableClass ? ` class=\"${tableClass}\"` : ""}>`];
  if (safeHeaders.length) {
    lines.push("  <thead>");
    lines.push("    <tr>");
    for (const header of safeHeaders) {
      lines.push(`      <th>${header}</th>`);
    }
    lines.push("    </tr>");
    lines.push("  </thead>");
  }
  lines.push("  <tbody>");
  for (const row of rows || []) {
    lines.push("    <tr>");
    for (const cell of row || []) {
      lines.push(`      <td>${render(cell || "")}</td>`);
    }
    lines.push("    </tr>");
  }
  lines.push("  </tbody>");
  lines.push("</table>");
  return lines;
}

const TAG_TO_HIGHLIGHT: Record<string, HighlightClass> = {
  treatment: "treatment",
  symptom: "symptom",
  clinical: "symptom",
  diagnostic: "diagnostic",
  lab: "diagnostic",
  labs: "diagnostic",
  imaging: "diagnostic",
  mechanism: "mechanism",
  "red flag": "symptom",
  "red flags": "symptom",
  cutoff: "cutoff",
  gene: "gene",
  enzyme: "enzyme",
  buzz: "buzz",
  buzzword: "buzz",
  buzzwords: "buzz",
  disease: "disease",
  dx: "disease",
  diagnosis: "disease",
  "risk factor": "disease",
  "risk factors": "disease",
  epidemiology: "disease",
  histology: "histology",
  pathophys: "histology",
};

const BUCKET_TO_HIGHLIGHT: Record<keyof StepABuckets, HighlightClass> = {
  dx: "disease",
  pathophys: "mechanism",
  clinical: "symptom",
  labs: "diagnostic",
  imaging: "diagnostic",
  treatment: "treatment",
  complications: "disease",
  risk_factors: "disease",
  epidemiology: "disease",
  red_flags: "symptom",
  buzzwords: "buzz",
};

function normalizeForMatch(value: string): string {
  return (value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function shouldIncludeHighlightPhrase(value: string): boolean {
  const trimmed = (value || "").trim();
  if (!trimmed) return false;
  if (trimmed.length < 3) return false;
  if (/^[0-9]+$/.test(trimmed)) return false;
  return true;
}

function resolveHighlightClassFromTags(tags: string[] | undefined | null): HighlightClass | null {
  for (const tag of normalizeList(tags)) {
    const normalized = normalizeForMatch(tag);
    if (!normalized) continue;
    const className = TAG_TO_HIGHLIGHT[normalized];
    if (className) return className;
  }
  return null;
}

function buildHighlightLexicon(stepA: StepAOutput): HighlightLexiconEntry[] {
  const entries: HighlightLexiconEntry[] = [];
  const seen = new Set<string>();
  const addEntry = (phrase: string, className: HighlightClass) => {
    if (!shouldIncludeHighlightPhrase(phrase)) return;
    const normalized = normalizeForMatch(phrase);
    if (!normalized) return;
    const key = `${className}:${normalized}`;
    if (seen.has(key)) return;
    seen.add(key);
    entries.push({ phrase, normalized, className });
  };

  const buckets = stepA?.buckets;
  if (buckets) {
    (Object.keys(buckets) as Array<keyof StepABuckets>).forEach((bucketName) => {
      const className = BUCKET_TO_HIGHLIGHT[bucketName];
      if (!className) return;
      for (const item of normalizeList(buckets[bucketName])) {
        addEntry(item, className);
      }
    });
  }

  for (const slide of normalizeList(stepA?.slides)) {
    for (const section of normalizeList(slide?.sections)) {
      for (const fact of normalizeList(section?.facts)) {
        const className = resolveHighlightClassFromTags(fact?.tags);
        if (!className) continue;
        addEntry(fact?.text || "", className);
      }
    }
  }

  entries.sort((a, b) => b.normalized.length - a.normalized.length);
  return entries;
}

function findHighlightClass(text: string, lexicon: HighlightLexiconEntry[]): HighlightClass | null {
  if (!lexicon.length) return null;
  const normalized = normalizeForMatch(text);
  if (!normalized) return null;
  for (const entry of lexicon) {
    if (normalized.includes(entry.normalized)) {
      return entry.className;
    }
  }
  return null;
}

function applyHighlights(text: string, lexicon: HighlightLexiconEntry[], tags?: string[]): string {
  const value = text || "";
  const tagClass = resolveHighlightClassFromTags(tags);
  const className = tagClass ?? findHighlightClass(value, lexicon);
  if (!className) {
    return escapeHtml(value);
  }
  return `<span class="hl ${className}">${escapeHtml(value)}</span>`;
}

/**
 * Render a maximal study guide HTML document from Step A/B/C outputs.
 *
 * @param opts - Rendering options including metadata and step outputs.
 * @returns A complete HTML document string.
 * @remarks Side effects: none (pure string builder).
 */
export function renderStudyGuideHtml(opts: {
  lectureTitle: string;
  buildUtc: string;
  stepA: StepAOutput;
  stepB: StepBOutput;
  stepC: StepCOutput;
  sourceTextBySlide?: Record<string, string>;
}): string {
  const lectureTitle = opts.lectureTitle || opts.stepA?.lecture_title || "Lecture";
  const buildUtc = opts.buildUtc || "1970-01-01T00:00:00Z";
  const slideCount = normalizeList(opts.stepA?.slides).length;
  const sourceTextBySlide = opts.sourceTextBySlide || {};
  const highlightLexicon = buildHighlightLexicon(opts.stepA);

  const lines: string[] = [];
  const push = (line: string) => lines.push(line);

  push("<!doctype html>");
  push("<html>");
  push("<head>");
  push("<meta charset=\"utf-8\" />");
  push("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />");
  push(`<title>${escapeHtml(lectureTitle)} Study Guide</title>`);
  push("<style>");
  push(BASE_STUDY_GUIDE_CSS);
  push("body { margin: 0; background: var(--bg-app); color: var(--text-primary); }");
  push(".sticky-header { z-index: 5; }");
  push("main.content { padding: 16px 20px 32px; }");
  push("section { margin-bottom: 28px; }");
  push("h1 { font-size: 1.35em; margin: 14px 0 8px; }");
  push("h2 { font-size: 1.1em; margin: 12px 0 6px; }");
  push("h3 { font-size: 1em; margin: 10px 0 6px; }");
  push(".muted { color: var(--text-secondary); }");
  push(".section-card { background: var(--bg-surface); border: 1px solid var(--border-subtle); border-radius: 10px; padding: 14px; box-shadow: var(--shadow-card); }");
  push(".slide-block { border-top: 1px solid var(--border-subtle); padding-top: 12px; margin-top: 12px; }");
  push(".slide-raw { white-space: pre-wrap; background: var(--bg-surface-alt); padding: 10px; border-radius: 6px; font-size: 0.95em; }");
  push(".tag-pill { display: inline-block; padding: 2px 6px; margin-right: 4px; border-radius: 999px; background: #eef2ff; font-size: 0.75em; font-weight: 600; color: var(--text-primary); }");
  push("</style>");
  push("</head>");
  push("<body>");
  push("<header class=\"sticky-header\">Study Guide - Maximal Build</header>");
  push("<nav class=\"toc\">");
  push("  <a href=\"#output-identity\">Output Identity</a>");
  push("  <a href=\"#document-map\">Document Map</a>");
  push("  <a href=\"#highlight-legend\">Highlight Legend</a>");
  push("  <a href=\"#high-yield-summary\">High-Yield Summary</a>");
  push("  <a href=\"#rapid-approach-summary\">Rapid-Approach Summary (Global)</a>");
  push("  <a href=\"#one-page-last-minute-review\">One-Page Last-Minute Review</a>");
  push("  <a href=\"#slide-by-slide-appendix\">Slide-by-Slide Appendix</a>");
  push("  <a href=\"#source-note-quality-assurance\">Source Note &amp; Quality Assurance</a>");
  push("</nav>");
  push("<main class=\"content\">");

  push("<section id=\"output-identity\" class=\"section-card\">");
  push("  <h1>Output Identity</h1>");
  push(`  <p>Lecture title: ${escapeHtml(lectureTitle)}</p>`);
  push(`  <p>Timestamp (UTC): ${escapeHtml(buildUtc)}</p>`);
  push(`  <p>Slide count: ${slideCount}</p>`);
  push("  <p>Content Hash: {{CONTENT_HASH}}</p>");
  push("</section>");

  push("<section id=\"document-map\" class=\"section-card\">");
  push("  <h1>Document Map</h1>");
  push("  <h2>Slide Outline</h2>");
  if (normalizeList(opts.stepA?.slides).length) {
    push("  <ul>");
    for (const slide of normalizeList(opts.stepA?.slides)) {
      const heading = slide?.sections?.[0]?.heading || "";
      const headingText = heading ? ` - ${escapeHtml(heading)}` : "";
      push(`    <li>Slide ${slide.n} (p.${slide.page})${headingText}</li>`);
    }
    push("  </ul>");
  } else {
    push("  <p class=\"muted\">No slides parsed.</p>");
  }

  push("</section>");

  push("<section id=\"highlight-legend\" class=\"section-card\">");
  push("  <h1>Highlight Legend</h1>");
  push(`  ${renderLegend()}`);
  push("</section>");

  push("<section id=\"high-yield-summary\" class=\"section-card\">");
  push("  <h1>High-Yield Summary</h1>");
  push(...renderList(normalizeList(opts.stepB?.high_yield_summary), "No high-yield summary items.", (item) => applyHighlights(item, highlightLexicon)));

  if (normalizeList(opts.stepB?.supplemental_glue).length) {
    push("  <h2>Supplemental Glue</h2>");
    push(...renderList(normalizeList(opts.stepB?.supplemental_glue), "", (item) => applyHighlights(item, highlightLexicon)));
  }

  if (normalizeList(opts.stepB?.compare_differential).length) {
    push("  <h2>Compare Differential</h2>");
    for (const topic of normalizeList(opts.stepB?.compare_differential)) {
      push(`  <h3>${escapeHtml(topic.topic || "Differential")}</h3>`);
      const rows = normalizeList(topic.rows).map(row => [row.dx1, row.dx2, row.how_to_tell]);
      push(...renderTable(["Dx 1", "Dx 2", "How to Tell"], rows, "compare", (cell) => applyHighlights(cell, highlightLexicon)));
    }
  }

  if (normalizeList(opts.stepB?.quant_cutoffs).length) {
    push("  <h2>Quant & Cutoffs</h2>");
    const rows = normalizeList(opts.stepB?.quant_cutoffs).map(item => [item.item, item.value, item.note]);
    push(...renderTable(["Item", "Value", "Note"], rows, "cutoff", (cell) => applyHighlights(cell, highlightLexicon)));
  }

  if (normalizeList(opts.stepB?.pitfalls).length) {
    push("  <h2>Pitfalls</h2>");
    push(...renderList(normalizeList(opts.stepB?.pitfalls), "", (item) => applyHighlights(item, highlightLexicon)));
  }

  if (normalizeList(opts.stepB?.glossary).length) {
    push("  <h2>Glossary</h2>");
    push("  <ul>");
    for (const item of normalizeList(opts.stepB?.glossary)) {
      const term = applyHighlights(item.term || "", highlightLexicon);
      const definition = applyHighlights(item.definition || "", highlightLexicon);
      push(`    <li><strong>${term}</strong>: ${definition}</li>`);
    }
    push("  </ul>");
  }
  push("</section>");

  push("<section id=\"rapid-approach-summary\" class=\"section-card\">");
  push("  <h1>Rapid-Approach Summary (Global)</h1>");
  if (normalizeList(opts.stepB?.rapid_approach_table).length) {
    const rows = normalizeList(opts.stepB?.rapid_approach_table).map(item => [item.clue, item.think_of, item.why, item.confirm]);
    push(
      ...renderTable(["Clue", "Think Of", "Why", "Confirm"], rows, "tri", (cell) => applyHighlights(cell, highlightLexicon)),
    );
  } else {
    push("  <p class=\"muted\">No rapid-approach table rows.</p>");
  }
  push("</section>");

  push("<section id=\"one-page-last-minute-review\" class=\"section-card\">");
  push("  <h1>One-Page Last-Minute Review</h1>");
  push(
    ...renderList(
      normalizeList(opts.stepB?.one_page_last_minute_review),
      "No review items.",
      (item) => applyHighlights(item, highlightLexicon),
    ),
  );
  push("</section>");

  push("<section id=\"slide-by-slide-appendix\" class=\"section-card\">");
  push("  <h1>Slide-by-Slide Appendix</h1>");
  if (!normalizeList(opts.stepA?.slides).length) {
    push("  <p class=\"muted\">No slide details available.</p>");
  }
  for (const slide of normalizeList(opts.stepA?.slides)) {
    const rawText = sourceTextBySlide[String(slide.n)] || "";
    push(`  <div class=\"slide-block\" id=\"slide-${slide.n}\">`);
    push(`    <h2>Slide ${slide.n} - (p.${slide.page})</h2>`);
    push(`    <div class=\"slide-raw\">${escapeHtml(rawText || "[NO TEXT]")}</div>`);

    if (normalizeList(slide.sections).length) {
      for (const section of normalizeList(slide.sections)) {
        push(`    <h3>${escapeHtml(section.heading || "Section")}</h3>`);
        if (normalizeList(section.facts).length) {
          push("    <ul>");
          for (const fact of normalizeList(section.facts)) {
            const tags = normalizeList(fact.tags).map(tag => `<span class=\"tag-pill\">${escapeHtml(tag)}</span>`).join("");
            const numbers = normalizeList(fact.numbers).map(num => `${escapeHtml(num.value)} ${escapeHtml(num.unit)}`.trim()).filter(Boolean).join(", ");
            const numbersText = numbers ? ` <span class=\"muted\">(${escapeHtml(numbers)})</span>` : "";
            const factText = applyHighlights(fact.text || "", highlightLexicon, fact.tags);
            push(`      <li>${tags}${factText}${numbersText}</li>`);
          }
          push("    </ul>");
        } else {
          push("    <p class=\"muted\">No facts parsed.</p>");
        }
      }
    }

    if (normalizeList(slide.tables).length) {
      push("    <h3>Tables</h3>");
      for (const table of normalizeList(slide.tables)) {
        if (table.caption) {
          push(`    <p class=\"muted\">${escapeHtml(table.caption)}</p>`);
        }
        push(...renderTable(table.headers || [], table.rows || []));
      }
    }

    push("  </div>");
  }
  push("</section>");

  push("<section id=\"source-note-quality-assurance\" class=\"section-card\">");
  push("  <h1>Source Note &amp; Quality Assurance</h1>");
  push(`  <p>Coverage confidence: ${escapeHtml(opts.stepC?.coverage_confidence || "")}</p>`);

  push("  <h2>Checks</h2>");
  const checks = opts.stepC?.checks;
  if (checks) {
    const rows = [
      ["has_high_yield_summary", String(Boolean(checks.has_high_yield_summary))],
      ["has_rapid_approach_table", String(Boolean(checks.has_rapid_approach_table))],
      ["has_one_page_review", String(Boolean(checks.has_one_page_review))],
      ["slide_count_stepA", String(checks.slide_count_stepA ?? "")],
      ["slide_count_rendered", String(checks.slide_count_rendered ?? "")],
    ];
    push(...renderTable(["Check", "Value"], rows));
  } else {
    push("  <p class=\"muted\">No checks reported.</p>");
  }

  push("  <h2>Unparsed Items</h2>");
  if (normalizeList(opts.stepC?.unparsed_items).length) {
    push("  <ul>");
    for (const item of normalizeList(opts.stepC?.unparsed_items)) {
      push(`    <li>Slide ${item.slide}: ${escapeHtml(item.note)}</li>`);
    }
    push("  </ul>");
  } else {
    push("  <p class=\"muted\">No unparsed items.</p>");
  }

  push("  <h2>Omissions</h2>");
  if (normalizeList(opts.stepC?.omissions).length) {
    push("  <ul>");
    for (const omission of normalizeList(opts.stepC?.omissions)) {
      push(`    <li>Slide ${omission.slide}: ${escapeHtml(omission.note)}</li>`);
    }
    push("  </ul>");
  } else {
    push("  <p class=\"muted\">No omissions flagged.</p>");
  }

  push("  <h2>Conflicts</h2>");
  if (normalizeList(opts.stepC?.conflicts).length) {
    push("  <ul>");
    for (const conflict of normalizeList(opts.stepC?.conflicts)) {
      push(`    <li>Slide ${conflict.slide}: ${escapeHtml(conflict.note)}</li>`);
    }
    push("  </ul>");
  } else {
    push("  <p class=\"muted\">No conflicts flagged.</p>");
  }
  push("</section>");

  push("</main>");
  push("</body>");
  push("</html>");

  return lines.join("\n");
}
