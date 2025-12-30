/**
 * Prompt templates for Step B study guide synthesis and normalization.
 *
 * Used by: `src/index.ts` to orchestrate multi-step study guide generation.
 *
 * Key exports:
 * - Step A2 prose synthesis, B1 draft, B2 normalization, and rewrite prompts.
 *
 * Assumptions:
 * - Each prompt expects strict output formatting (JSON or plain text).
 */
/** Step A2 prose synthesis from Step A JSON. */
export const STUDY_GUIDE_STEP_A2_SYNTHESIS_PROMPT = String.raw`You are a medical study-guide synthesis engine.
Input is Step A JSON containing extracted facts. Output MUST be plain prose (no JSON, no schema, no code fences).

Goal:
- Produce exam-oriented prose that preserves all facts and nuance.
- Do NOT compress or summarize away specifics.

Rules:
- Use clear headings and short paragraphs or bullet clusters.
- Do not invent facts or add external knowledge.
- Preserve numbers, thresholds, and table relationships; describe any tables explicitly.
- Keep all clinically important details; do not omit edge cases.
- Do NOT output JSON or markdown fences.

LECTURE_TITLE:
{{LECTURE_TITLE}}

STEP_A_JSON:
{{STEP_A_JSON}}
`;

/** Step B1 draft authoring from Step A2 prose. */
export const STUDY_GUIDE_STEP_B1_AUTHOR_PROMPT = String.raw`You are a medical study-guide authoring engine.
Input is Step A2 prose. Draft the full guide with section headers and tables written naturally.
Output MUST be plain text with headings and bullets. NO JSON, NO schema, NO code fences.

Rules:
- Use ONLY the supplied prose; do not add external facts.
- If content is sparse, reuse phrasing from the prose; do not invent.
- Keep bullets concise and exam-forward.
- Write tables in plain text; use consistent row formats.

Required sections (use these headings):
1) High-Yield Summary (8-12 bullets, aim <= 16 words each)
2) Rapid Approach Table (10-18 rows; format each row as "Clue | Think of | Why | Confirm")
3) One-Page Last-Minute Review (12-18 bullets, aim <= 14 words each)
4) Compare Differential (2-4 topics; each topic 4-7 rows of "Dx1 vs Dx2 - how to tell")

Optional sections (only if supported by the prose):
- Quant Cutoffs (rows: "Item - Value - Note")
- Pitfalls (bullets)
- Glossary (rows: "Term - Definition")
- Supplemental Glue (bullets; only if needed to bridge concepts)

STEP_A2_PROSE:
{{STEP_A2_PROSE}}
`;

/** Normalize the Step B1 draft into the required JSON schema. */
export const STUDY_GUIDE_STEP_B2_NORMALIZE_PROMPT = String.raw`You are a schema-normalization engine.
Rewrite the draft below into EXACTLY this JSON schema. Output ONLY a single JSON object.

Rules:
- Use ONLY the draft content; do not add external facts.
- Preserve wording when possible; trim to fit limits.
- If a section is unsupported or missing, return an empty array for that field.
- Do not include any extra keys or commentary.

Schema (return exactly these keys):
{
  "high_yield_summary": ["string"],
  "rapid_approach_table": [
    {"clue":"string","think_of":"string","why":"string","confirm":"string"}
  ],
  "one_page_last_minute_review": ["string"],
  "compare_differential": [
    {"topic":"string","rows":[{"dx1":"string","dx2":"string","how_to_tell":"string"}]}
  ],
  "quant_cutoffs": [{"item":"string","value":"string","note":"string"}],
  "pitfalls": ["string"],
  "glossary": [{"term":"string","definition":"string"}],
  "supplemental_glue": ["string"]
}

Constraints:
- high_yield_summary: 8-12 bullets, each <= 16 words.
- one_page_last_minute_review: 12-18 bullets, each <= 14 words.
- rapid_approach_table: 10-18 rows; clue <= 10 words; think_of <= 6 words; why <= 14 words; confirm <= 10 words.
- compare_differential: 2-4 topics; each topic has 4-7 rows; how_to_tell <= 18 words.
- supplemental_glue: max 10 items, each <= 14 words.

DRAFT:
{{STEP_B1_DRAFT}}
`;

/** Direct Step B synthesis prompt (JSON) from Step A output. */
export const STUDY_GUIDE_STEP_B_ENHANCED_PROMPT = String.raw`You are a medical study-guide synthesis engine.
Return ONLY a single JSON object. No markdown. No commentary. No trailing commas.

Use ONLY the provided Step A JSON. Do not add external knowledge.
Keep items concise and exam-forward.

Schema:
{
  "high_yield_summary": ["string"],
  "rapid_approach_table": [
    {"clue":"string","think_of":"string","why":"string","confirm":"string"}
  ],
  "one_page_last_minute_review": ["string"],
  "compare_differential": [
    {"topic":"string","rows":[{"dx1":"string","dx2":"string","how_to_tell":"string"}]}
  ],
  "quant_cutoffs": [{"item":"string","value":"string","note":"string"}],
  "pitfalls": ["string"],
  "glossary": [{"term":"string","definition":"string"}],
  "supplemental_glue": ["string"]
}

Minimums:
- high_yield_summary >= 8 (each <= 16 words)
- rapid_approach_table >= 10 rows (clue <= 10 words; think_of <= 6; why <= 14; confirm <= 10)
- one_page_last_minute_review >= 12 (each <= 14 words)

Constraints:
- compare_differential: 2-4 topics; each topic 4-7 rows; how_to_tell <= 18 words.
- supplemental_glue: max 10 items, each <= 14 words.
- quant_cutoffs, pitfalls, glossary: include only if supported; otherwise [].

STEP_A_JSON:
{{STEP_A_JSON}}
`;

/** Planning prompt for Step B outlines and constraints. */
export const STUDY_GUIDE_STEP_B_PLAN_PROMPT = String.raw`You are a medical study-guide planner.
Return ONLY a single JSON object. No markdown. No commentary. No trailing commas.

Use ONLY the provided Step A plan JSON to design Step B structure and counts.
Keep counts within validator limits:
- high_yield_summary: 8-12
- one_page_last_minute_review: 12-18
- rapid_approach_table_rows: 10-18
- compare_topics: 2-4
- compare_rows_per_topic: 4-7

Output schema:
{
  "selected_exam_atoms": ["string"],
  "section_counts": {
    "high_yield_summary": 10,
    "one_page_last_minute_review": 16,
    "rapid_approach_table_rows": 14,
    "compare_topics": 3,
    "compare_rows_per_topic": 5
  },
  "compare_topics": ["string"],
  "atom_to_section_map": [{"atom":"string","section":"high_yield_summary|one_page_last_minute_review|rapid_approach_table|compare_differential"}],
  "warnings": ["string"]
}

Rules:
- selected_exam_atoms: pick 12-18 high-yield atoms from Step A (exam_atoms preferred).
- compare_topics: prefer discriminator topics; fall back to dx bucket items.
- atom_to_section_map should include each selected atom once.
- warnings: [] if none; include short strings for sparsity or weak coverage.

STEP_A_JSON:
{{STEP_A_JSON}}
`;

/** Outline prompt used before drafting Step B content. */
export const STUDY_GUIDE_STEP_B1_OUTLINE_PROMPT = String.raw`You are a medical study-guide outline engine.
Output a plain-text outline only (no JSON, no markdown fences).

Use the plan to draft a section-by-section outline. Keep wording exam-forward.
Use these exact section headers:
HIGH_YIELD_SUMMARY
RAPID_APPROACH_TABLE
ONE_PAGE_LAST_MINUTE_REVIEW
COMPARE_DIFFERENTIAL
QUANT_CUTOFFS (optional)
PITFALLS (optional)
GLOSSARY (optional)
SUPPLEMENTAL_GLUE (optional)

Formatting rules:
- Use bullets with "-" for items.
- For table rows, use "Clue | Think of | Why | Confirm".
- For compare differential rows, use "Dx1 | Dx2 | How to tell".

STEP_A_PLAN_JSON:
{{STEP_A_PLAN_JSON}}

SECTION_COUNTS_JSON:
{{SECTION_COUNTS_JSON}}

COMPARE_TOPICS_JSON:
{{COMPARE_TOPICS_JSON}}
`;

/** Pack Step B outline + draft into the JSON schema. */
export const STUDY_GUIDE_STEP_B2_PACK_JSON_PROMPT = String.raw`You are a medical study-guide synthesis engine.
Return ONLY a single JSON object. No markdown. No commentary. No trailing commas.

Use ONLY Step A facts. Follow the outline to structure the output.
Keep items concise, exam-forward, and within word limits.

Schema:
{
  "high_yield_summary": ["string"],
  "rapid_approach_table": [
    {"clue":"string","think_of":"string","why":"string","confirm":"string"}
  ],
  "one_page_last_minute_review": ["string"],
  "compare_differential": [
    {"topic":"string","rows":[{"dx1":"string","dx2":"string","how_to_tell":"string"}]}
  ],
  "quant_cutoffs": [{"item":"string","value":"string","note":"string"}],
  "pitfalls": ["string"],
  "glossary": [{"term":"string","definition":"string"}],
  "supplemental_glue": ["string"]
}

Constraints:
- high_yield_summary: 8-12 bullets, each <= 16 words.
- one_page_last_minute_review: 12-18 bullets, each <= 14 words.
- rapid_approach_table: 10-18 rows; clue <= 10 words; think_of <= 6 words; why <= 14 words; confirm <= 10 words.
- compare_differential: 2-4 topics; each topic has 4-7 rows; how_to_tell <= 18 words.
- supplemental_glue: max 10 items, each <= 14 words.
- quant_cutoffs, pitfalls, glossary: include only if supported; otherwise [].

STEP_A_JSON:
{{STEP_A_JSON}}

STEP_B1_OUTLINE:
{{STEP_B1_OUTLINE}}
`;

/** Rewrite Step B JSON for quality/consistency. */
export const STUDY_GUIDE_STEP_B2_REWRITE_JSON_PROMPT = String.raw`You are a medical study-guide rewrite engine.
Fix the draft to satisfy all validation failures. Return ONLY a single JSON object.

Rules:
- Use ONLY Step A facts.
- Preserve good content; fix counts, empty fields, and word limits.
- Resolve each failure explicitly.
- No extra keys or commentary.

Schema and constraints are identical to the Step B JSON schema.

STEP_A_JSON:
{{STEP_A_JSON}}

STEP_B1_OUTLINE:
{{STEP_B1_OUTLINE}}

STEP_B_DRAFT_JSON:
{{STEP_B_DRAFT_JSON}}

STEP_B_FAILURES_JSON:
{{STEP_B_FAILURES_JSON}}
`;

/** QC rewrite prompt to fix validation issues in Step B output. */
export const STUDY_GUIDE_STEP_B_QC_REWRITE_PROMPT = String.raw`You are a medical study-guide QC rewrite engine.
Fix the draft to satisfy all validation failures. Return ONLY a single JSON object.

Rules:
- Use ONLY Step A facts.
- Preserve good content; fix counts, empty fields, and word limits.
- Resolve each failure explicitly.
- No extra keys or commentary.

STEP_A_JSON:
{{STEP_A_JSON}}

STEP_B_DRAFT_JSON:
{{STEP_B_DRAFT_JSON}}

STEP_B_FAILURES_JSON:
{{STEP_B_FAILURES_JSON}}
`;

/** Draft prompt used for Step B authoring when bypassing Step A2. */
export const STUDY_GUIDE_STEP_B_DRAFT_PROMPT = String.raw`You are a medical study-guide authoring engine.
Return ONLY a single JSON object. No markdown. No commentary. No trailing commas.

Use ONLY Step A facts and the plan. Meet minimum synthesis counts.

Schema:
{
  "high_yield_summary": ["string"],
  "rapid_approach_table": [
    {"clue":"string","think_of":"string","why":"string","confirm":"string"}
  ],
  "one_page_last_minute_review": ["string"],
  "compare_differential": [
    {"topic":"string","rows":[{"dx1":"string","dx2":"string","how_to_tell":"string"}]}
  ],
  "quant_cutoffs": [{"item":"string","value":"string","note":"string"}],
  "pitfalls": ["string"],
  "glossary": [{"term":"string","definition":"string"}],
  "supplemental_glue": ["string"]
}

Minimums:
- high_yield_summary >= 8 (<= 16 words each)
- rapid_approach_table >= 10 rows (each field within word limits)
- one_page_last_minute_review >= 12 (<= 14 words each)

STEP_A_JSON:
{{STEP_A_JSON}}

STEP_B_PLAN_JSON:
{{STEP_B_PLAN_JSON}}
`;

/** Rewrite prompt for Step B synthesis output (post-validation). */
export const STUDY_GUIDE_STEP_B_SYNTHESIS_REWRITE_PROMPT = String.raw`You are a medical study-guide rewrite engine.
Repair the draft to meet synthesis minimums. Return ONLY a single JSON object.

Rules:
- Use ONLY Step A facts.
- Keep wording concise and exam-forward.
- Ensure required arrays meet minimum counts.
- No extra keys or commentary.

STEP_A_JSON:
{{STEP_A_JSON}}

STEP_B_PLAN_JSON:
{{STEP_B_PLAN_JSON}}

STEP_B_DRAFT_JSON:
{{STEP_B_DRAFT_JSON}}

STEP_B_FAILURES_JSON:
{{STEP_B_FAILURES_JSON}}
`;
