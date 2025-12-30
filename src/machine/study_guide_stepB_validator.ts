/**
 * Validation utilities for Step B study guide synthesis outputs.
 *
 * Used by: `src/index.ts` and tests to enforce structure, coverage, and
 * non-redundancy constraints on the Step B JSON payload.
 *
 * Key exports:
 * - Failure enums/types and `STEP_B_VALIDATION_LIMITS` for consistent rules.
 * - `validateStepB` for full validation against Step A coverage.
 * - `validateSynthesis`/`validateSynthesisOrThrow` for minimum completeness.
 *
 * Assumptions:
 * - Inputs follow the Step A/B schema from `render_study_guide_html.ts`.
 */
import type { StepAOutput, StepBOutput } from "./render_study_guide_html";

/**
 * Codes for structural/content failures in Step B output.
 */
export type StepBFailureCode =
  | "BULLET_TOO_LONG"
  | "TOO_MANY_BULLETS"
  | "TOO_FEW_BULLETS"
  | "REDUNDANT_BULLETS"
  | "HIGH_NGRAM_OVERLAP"
  | "LOW_COVERAGE"
  | "TABLE_ROW_INVALID"
  | "GLUE_RULE_VIOLATION";

/**
 * Structured validation failure for Step B checks.
 */
export type StepBValidationFailure = {
  code: StepBFailureCode;
  message: string;
  path?: string;
};

/**
 * Codes for minimum synthesis completeness checks.
 */
export type StepBSynthesisFailureCode =
  | "SYNTHESIS_MISSING"
  | "SYNTHESIS_TOO_FEW"
  | "SYNTHESIS_EMPTY";

/**
 * Structured failure when synthesis output is missing or too sparse.
 */
export type StepBSynthesisFailure = {
  code: StepBSynthesisFailureCode;
  message: string;
  path?: string;
};

/**
 * Validation thresholds for Step B output formatting and coverage.
 */
export const STEP_B_VALIDATION_LIMITS = {
  high_yield_summary: { min: 8, max: 12, max_words: 16 },
  one_page_last_minute_review: { min: 12, max: 18, max_words: 14 },
  rapid_approach_table: {
    min_rows: 10,
    max_rows: 18,
    clue_max_words: 10,
    think_of_max_words: 6,
    why_max_words: 14,
    confirm_max_words: 10,
  },
  compare_differential: { min_topics: 2, max_topics: 4, min_rows: 4, max_rows: 7, how_to_tell_max_words: 18 },
  supplemental_glue: { max_items: 10, max_words: 14 },
  coverage: { min_ratio: 0.7, atom_token_overlap: 0.6 },
  redundancy: { trigram_repeat_limit: 3 },
} as const;

const SYNTHESIS_MINIMUMS = {
  high_yield_summary: 8,
  rapid_approach_table: 10,
  one_page_last_minute_review: 12,
} as const;

const SYNTHESIS_NONEMPTY_RATIO = 0.7;

const STOPWORDS = new Set([
  "a",
  "an",
  "and",
  "are",
  "as",
  "at",
  "be",
  "by",
  "for",
  "from",
  "in",
  "is",
  "it",
  "of",
  "on",
  "or",
  "the",
  "to",
  "with",
  "without",
]);

const WORD_RE = /[a-z0-9]+/gi;

function toWords(text: string): string[] {
  return (text || "").match(WORD_RE) || [];
}

function wordCount(text: string): number {
  return toWords(text).length;
}

function nonEmptyCount(items: string[]): number {
  let count = 0;
  for (const item of items) {
    if ((item || "").trim()) count += 1;
  }
  return count;
}

function ratio(part: number, total: number): number {
  if (!total) return 0;
  return part / total;
}

function normalizeText(text: string): string {
  return (text || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokensForCoverage(text: string): string[] {
  const normalized = normalizeText(text);
  if (!normalized) return [];
  return normalized
    .split(" ")
    .filter(token => token.length >= 2 && !STOPWORDS.has(token));
}

function tokensForNgrams(text: string): string[] {
  const normalized = normalizeText(text);
  if (!normalized) return [];
  return normalized.split(" ").filter(Boolean);
}

function collectStepBStrings(stepB: StepBOutput): string[] {
  const strings: string[] = [];
  for (const item of stepB.high_yield_summary || []) strings.push(item);
  for (const item of stepB.one_page_last_minute_review || []) strings.push(item);
  for (const item of stepB.pitfalls || []) strings.push(item);
  for (const item of stepB.supplemental_glue || []) strings.push(item);
  for (const row of stepB.rapid_approach_table || []) {
    strings.push(row?.clue || "", row?.think_of || "", row?.why || "", row?.confirm || "");
  }
  for (const topic of stepB.compare_differential || []) {
    strings.push(topic?.topic || "");
    for (const row of topic?.rows || []) {
      strings.push(row?.dx1 || "", row?.dx2 || "", row?.how_to_tell || "");
    }
  }
  for (const item of stepB.quant_cutoffs || []) {
    strings.push(item?.item || "", item?.value || "", item?.note || "");
  }
  for (const item of stepB.glossary || []) {
    strings.push(item?.term || "", item?.definition || "");
  }
  return strings.filter(value => value && value.trim());
}

function collectStepAStrings(stepA: StepAOutput): string[] {
  const strings: string[] = [];
  for (const item of stepA.raw_facts || []) strings.push(item);
  for (const item of stepA.exam_atoms || []) strings.push(item);
  for (const item of stepA.buckets?.dx || []) strings.push(item);
  for (const item of stepA.buckets?.pathophys || []) strings.push(item);
  for (const item of stepA.buckets?.clinical || []) strings.push(item);
  for (const item of stepA.buckets?.labs || []) strings.push(item);
  for (const item of stepA.buckets?.imaging || []) strings.push(item);
  for (const item of stepA.buckets?.treatment || []) strings.push(item);
  for (const item of stepA.buckets?.complications || []) strings.push(item);
  for (const item of stepA.buckets?.risk_factors || []) strings.push(item);
  for (const item of stepA.buckets?.epidemiology || []) strings.push(item);
  for (const item of stepA.buckets?.red_flags || []) strings.push(item);
  for (const item of stepA.buckets?.buzzwords || []) strings.push(item);
  for (const item of stepA.discriminators || []) {
    if (item?.topic) strings.push(item.topic);
    for (const signal of item?.signals || []) strings.push(signal);
    for (const pitfall of item?.pitfalls || []) strings.push(pitfall);
  }
  for (const [abbr, expansion] of Object.entries(stepA.abbrev_map || {})) {
    strings.push(abbr);
    strings.push(expansion);
  }
  for (const slide of stepA.slides || []) {
    for (const section of slide.sections || []) {
      for (const fact of section.facts || []) {
        if (fact?.text) strings.push(fact.text);
      }
    }
    for (const table of slide.tables || []) {
      if (table?.caption) strings.push(table.caption);
      for (const header of table?.headers || []) strings.push(header);
      for (const row of table?.rows || []) {
        for (const cell of row || []) strings.push(cell);
      }
    }
  }
  return strings.filter(value => value && value.trim());
}

function buildTokenSet(strings: string[]): Set<string> {
  const tokens = new Set<string>();
  for (const value of strings) {
    for (const token of tokensForCoverage(value)) {
      tokens.add(token);
    }
  }
  return tokens;
}

function includesAbbrev(glue: string, abbrevMap: Record<string, string>): boolean {
  const normalizedGlue = normalizeText(glue);
  if (!normalizedGlue) return false;
  for (const [abbr, expansion] of Object.entries(abbrevMap || {})) {
    const normalizedAbbr = normalizeText(abbr);
    const normalizedExpansion = normalizeText(expansion);
    if (normalizedAbbr && normalizedGlue.includes(normalizedAbbr)) return true;
    if (normalizedExpansion && normalizedGlue.includes(normalizedExpansion)) return true;
  }
  return false;
}

/**
 * Validate Step B output against Step A for coverage and formatting.
 *
 * @param stepA - Step A extraction output used as coverage baseline.
 * @param stepB - Step B synthesis output to validate.
 * @param selectedExamAtoms - Optional subset of Step A atoms to enforce.
 * @returns List of validation failures (empty when valid).
 */
export function validateStepB(
  stepA: StepAOutput,
  stepB: StepBOutput,
  selectedExamAtoms?: string[],
): StepBValidationFailure[] {
  const failures: StepBValidationFailure[] = [];
  const addFailure = (code: StepBFailureCode, message: string, path?: string) => {
    failures.push({ code, message, path });
  };

  const highYield = Array.isArray(stepB.high_yield_summary) ? stepB.high_yield_summary : [];
  const onePage = Array.isArray(stepB.one_page_last_minute_review) ? stepB.one_page_last_minute_review : [];
  const rapid = Array.isArray(stepB.rapid_approach_table) ? stepB.rapid_approach_table : [];
  const compare = Array.isArray(stepB.compare_differential) ? stepB.compare_differential : [];
  const glue = Array.isArray(stepB.supplemental_glue) ? stepB.supplemental_glue : [];

  if (highYield.length < STEP_B_VALIDATION_LIMITS.high_yield_summary.min) {
    addFailure("TOO_FEW_BULLETS", "High-yield summary has too few bullets.", "high_yield_summary");
  }
  if (highYield.length > STEP_B_VALIDATION_LIMITS.high_yield_summary.max) {
    addFailure("TOO_MANY_BULLETS", "High-yield summary has too many bullets.", "high_yield_summary");
  }
  highYield.forEach((item, idx) => {
    if (wordCount(item) > STEP_B_VALIDATION_LIMITS.high_yield_summary.max_words) {
      addFailure("BULLET_TOO_LONG", "High-yield summary bullet exceeds word limit.", `high_yield_summary[${idx}]`);
    }
  });

  if (onePage.length < STEP_B_VALIDATION_LIMITS.one_page_last_minute_review.min) {
    addFailure("TOO_FEW_BULLETS", "One-page review has too few bullets.", "one_page_last_minute_review");
  }
  if (onePage.length > STEP_B_VALIDATION_LIMITS.one_page_last_minute_review.max) {
    addFailure("TOO_MANY_BULLETS", "One-page review has too many bullets.", "one_page_last_minute_review");
  }
  onePage.forEach((item, idx) => {
    if (wordCount(item) > STEP_B_VALIDATION_LIMITS.one_page_last_minute_review.max_words) {
      addFailure("BULLET_TOO_LONG", "One-page review bullet exceeds word limit.", `one_page_last_minute_review[${idx}]`);
    }
  });

  if (rapid.length < STEP_B_VALIDATION_LIMITS.rapid_approach_table.min_rows) {
    addFailure("TOO_FEW_BULLETS", "Rapid-approach table has too few rows.", "rapid_approach_table");
  }
  if (rapid.length > STEP_B_VALIDATION_LIMITS.rapid_approach_table.max_rows) {
    addFailure("TOO_MANY_BULLETS", "Rapid-approach table has too many rows.", "rapid_approach_table");
  }
  rapid.forEach((row, idx) => {
    const clue = (row?.clue || "").trim();
    const thinkOf = (row?.think_of || "").trim();
    const why = (row?.why || "").trim();
    const confirm = (row?.confirm || "").trim();
    if (!clue || !thinkOf || !why || !confirm) {
      addFailure("TABLE_ROW_INVALID", "Rapid-approach row missing required fields.", `rapid_approach_table[${idx}]`);
      return;
    }
    if (wordCount(clue) > STEP_B_VALIDATION_LIMITS.rapid_approach_table.clue_max_words) {
      addFailure("BULLET_TOO_LONG", "Rapid-approach clue exceeds word limit.", `rapid_approach_table[${idx}].clue`);
    }
    if (wordCount(thinkOf) > STEP_B_VALIDATION_LIMITS.rapid_approach_table.think_of_max_words) {
      addFailure("BULLET_TOO_LONG", "Rapid-approach think_of exceeds word limit.", `rapid_approach_table[${idx}].think_of`);
    }
    if (wordCount(why) > STEP_B_VALIDATION_LIMITS.rapid_approach_table.why_max_words) {
      addFailure("BULLET_TOO_LONG", "Rapid-approach why exceeds word limit.", `rapid_approach_table[${idx}].why`);
    }
    if (wordCount(confirm) > STEP_B_VALIDATION_LIMITS.rapid_approach_table.confirm_max_words) {
      addFailure("BULLET_TOO_LONG", "Rapid-approach confirm exceeds word limit.", `rapid_approach_table[${idx}].confirm`);
    }
  });

  if (compare.length < STEP_B_VALIDATION_LIMITS.compare_differential.min_topics) {
    addFailure("TOO_FEW_BULLETS", "Compare differential has too few topics.", "compare_differential");
  }
  if (compare.length > STEP_B_VALIDATION_LIMITS.compare_differential.max_topics) {
    addFailure("TOO_MANY_BULLETS", "Compare differential has too many topics.", "compare_differential");
  }
  compare.forEach((topic, topicIdx) => {
    const topicLabel = (topic?.topic || "").trim();
    if (!topicLabel) {
      addFailure("TABLE_ROW_INVALID", "Compare differential topic is empty.", `compare_differential[${topicIdx}].topic`);
    }
    const rows = Array.isArray(topic?.rows) ? topic.rows : [];
    if (rows.length < STEP_B_VALIDATION_LIMITS.compare_differential.min_rows) {
      addFailure("TOO_FEW_BULLETS", "Compare differential topic has too few rows.", `compare_differential[${topicIdx}].rows`);
    }
    if (rows.length > STEP_B_VALIDATION_LIMITS.compare_differential.max_rows) {
      addFailure("TOO_MANY_BULLETS", "Compare differential topic has too many rows.", `compare_differential[${topicIdx}].rows`);
    }
    rows.forEach((row, rowIdx) => {
      const dx1 = (row?.dx1 || "").trim();
      const dx2 = (row?.dx2 || "").trim();
      const how = (row?.how_to_tell || "").trim();
      if (!dx1 || !dx2 || !how) {
        addFailure(
          "TABLE_ROW_INVALID",
          "Compare differential row missing required fields.",
          `compare_differential[${topicIdx}].rows[${rowIdx}]`,
        );
        return;
      }
      if (wordCount(how) > STEP_B_VALIDATION_LIMITS.compare_differential.how_to_tell_max_words) {
        addFailure(
          "BULLET_TOO_LONG",
          "Compare differential how_to_tell exceeds word limit.",
          `compare_differential[${topicIdx}].rows[${rowIdx}].how_to_tell`,
        );
      }
    });
  });

  if (glue.length > STEP_B_VALIDATION_LIMITS.supplemental_glue.max_items) {
    addFailure("TOO_MANY_BULLETS", "Supplemental glue has too many items.", "supplemental_glue");
  }
  glue.forEach((item, idx) => {
    if (wordCount(item) > STEP_B_VALIDATION_LIMITS.supplemental_glue.max_words) {
      addFailure("BULLET_TOO_LONG", "Supplemental glue exceeds word limit.", `supplemental_glue[${idx}]`);
    }
  });

  const normalizedBullets = collectStepBStrings(stepB).map(value => normalizeText(value)).filter(Boolean);
  const duplicateSet = new Set<string>();
  const duplicateFound = new Set<string>();
  for (const value of normalizedBullets) {
    if (duplicateSet.has(value)) {
      duplicateFound.add(value);
    } else {
      duplicateSet.add(value);
    }
  }
  if (duplicateFound.size) {
    addFailure("REDUNDANT_BULLETS", "Duplicate bullets detected across Step B output.");
  }

  const trigramCounts = new Map<string, number>();
  for (const value of normalizedBullets) {
    const tokens = tokensForNgrams(value);
    if (tokens.length < 3) continue;
    for (let i = 0; i <= tokens.length - 3; i += 1) {
      const t1 = tokens[i];
      const t2 = tokens[i + 1];
      const t3 = tokens[i + 2];
      if (STOPWORDS.has(t1) && STOPWORDS.has(t2) && STOPWORDS.has(t3)) continue;
      const trigram = `${t1} ${t2} ${t3}`;
      trigramCounts.set(trigram, (trigramCounts.get(trigram) || 0) + 1);
    }
  }
  const repeated = Array.from(trigramCounts.values()).some(count => count > STEP_B_VALIDATION_LIMITS.redundancy.trigram_repeat_limit);
  if (repeated) {
    addFailure("HIGH_NGRAM_OVERLAP", "Repeated trigrams detected across Step B output.");
  }

  const selectedAtoms = (selectedExamAtoms || []).length ? selectedExamAtoms || [] : stepA.exam_atoms || [];
  const stepBTokens = buildTokenSet(collectStepBStrings(stepB));
  let coveredAtoms = 0;
  let atomCount = 0;
  for (const atom of selectedAtoms || []) {
    const tokens = tokensForCoverage(atom);
    if (!tokens.length) continue;
    atomCount += 1;
    const matched = tokens.filter(token => stepBTokens.has(token)).length;
    const requiredRatio = STEP_B_VALIDATION_LIMITS.coverage.atom_token_overlap;
    const ratio = matched / tokens.length;
    const covered = tokens.length <= 2 ? ratio === 1 : ratio >= requiredRatio;
    if (covered) coveredAtoms += 1;
  }
  if (atomCount > 0) {
    const coverage = coveredAtoms / atomCount;
    if (coverage < STEP_B_VALIDATION_LIMITS.coverage.min_ratio) {
      addFailure(
        "LOW_COVERAGE",
        `Exam atom coverage ${Math.round(coverage * 100)}% is below target.`,
        "coverage",
      );
    }
  }

  const stepATokens = buildTokenSet(collectStepAStrings(stepA));
  if (glue.length) {
    for (let idx = 0; idx < glue.length; idx += 1) {
      const item = glue[idx] || "";
      if (!item.trim()) continue;
      const hasAbbrev = includesAbbrev(item, stepA.abbrev_map || {});
      const tokens = tokensForCoverage(item);
      const overlap = tokens.length
        ? tokens.filter(token => stepATokens.has(token)).length / tokens.length
        : 0;
      if (!hasAbbrev && overlap < STEP_B_VALIDATION_LIMITS.coverage.atom_token_overlap) {
        addFailure(
          "GLUE_RULE_VIOLATION",
          "Supplemental glue contains content not supported by Step A.",
          `supplemental_glue[${idx}]`,
        );
      }
    }
  }

  return failures;
}

/**
 * Validate Step B output for minimum synthesis completeness only.
 *
 * @param stepB - Step B synthesis output to validate.
 * @returns List of synthesis failures (empty when valid).
 */
export function validateSynthesis(stepB: StepBOutput): StepBSynthesisFailure[] {
  const failures: StepBSynthesisFailure[] = [];
  const addFailure = (code: StepBSynthesisFailureCode, message: string, path?: string) => {
    failures.push({ code, message, path });
  };

  const highYield = Array.isArray(stepB.high_yield_summary) ? stepB.high_yield_summary : [];
  const onePage = Array.isArray(stepB.one_page_last_minute_review) ? stepB.one_page_last_minute_review : [];
  const rapid = Array.isArray(stepB.rapid_approach_table) ? stepB.rapid_approach_table : [];

  if (!Array.isArray(stepB.high_yield_summary)) {
    addFailure("SYNTHESIS_MISSING", "High-yield summary is missing.", "high_yield_summary");
  } else if (highYield.length < SYNTHESIS_MINIMUMS.high_yield_summary) {
    addFailure("SYNTHESIS_TOO_FEW", "High-yield summary below minimum count.", "high_yield_summary");
  }
  const highYieldNonEmpty = nonEmptyCount(highYield);
  if (highYield.length && ratio(highYieldNonEmpty, highYield.length) < SYNTHESIS_NONEMPTY_RATIO) {
    addFailure("SYNTHESIS_EMPTY", "High-yield summary has mostly empty items.", "high_yield_summary");
  }
  if (highYieldNonEmpty && highYieldNonEmpty < SYNTHESIS_MINIMUMS.high_yield_summary) {
    addFailure("SYNTHESIS_TOO_FEW", "High-yield summary has too few non-empty items.", "high_yield_summary");
  }

  if (!Array.isArray(stepB.one_page_last_minute_review)) {
    addFailure("SYNTHESIS_MISSING", "One-page review is missing.", "one_page_last_minute_review");
  } else if (onePage.length < SYNTHESIS_MINIMUMS.one_page_last_minute_review) {
    addFailure("SYNTHESIS_TOO_FEW", "One-page review below minimum count.", "one_page_last_minute_review");
  }
  const onePageNonEmpty = nonEmptyCount(onePage);
  if (onePage.length && ratio(onePageNonEmpty, onePage.length) < SYNTHESIS_NONEMPTY_RATIO) {
    addFailure("SYNTHESIS_EMPTY", "One-page review has mostly empty items.", "one_page_last_minute_review");
  }
  if (onePageNonEmpty && onePageNonEmpty < SYNTHESIS_MINIMUMS.one_page_last_minute_review) {
    addFailure("SYNTHESIS_TOO_FEW", "One-page review has too few non-empty items.", "one_page_last_minute_review");
  }

  if (!Array.isArray(stepB.rapid_approach_table)) {
    addFailure("SYNTHESIS_MISSING", "Rapid-approach table is missing.", "rapid_approach_table");
  } else if (rapid.length < SYNTHESIS_MINIMUMS.rapid_approach_table) {
    addFailure("SYNTHESIS_TOO_FEW", "Rapid-approach table below minimum rows.", "rapid_approach_table");
  }
  const validRapidRows = rapid.filter((row) => {
    const clue = (row?.clue || "").trim();
    const thinkOf = (row?.think_of || "").trim();
    const why = (row?.why || "").trim();
    const confirm = (row?.confirm || "").trim();
    return clue && thinkOf && why && confirm;
  }).length;
  if (rapid.length && ratio(validRapidRows, rapid.length) < SYNTHESIS_NONEMPTY_RATIO) {
    addFailure("SYNTHESIS_EMPTY", "Rapid-approach table has mostly empty rows.", "rapid_approach_table");
  }
  if (validRapidRows && validRapidRows < SYNTHESIS_MINIMUMS.rapid_approach_table) {
    addFailure("SYNTHESIS_TOO_FEW", "Rapid-approach table has too few complete rows.", "rapid_approach_table");
  }

  return failures;
}

/**
 * Validate synthesis output and throw when failures are present.
 *
 * @param stepB - Step B synthesis output to validate.
 * @returns The failure list (empty when valid).
 * @throws Error when validation failures are detected.
 */
export function validateSynthesisOrThrow(stepB: StepBOutput): StepBSynthesisFailure[] {
  const failures = validateSynthesis(stepB);
  if (!failures.length) return failures;
  const codes = Array.from(new Set(failures.map(item => item.code))).join(", ");
  throw new Error(`Step B synthesis failed: ${codes}`);
}
