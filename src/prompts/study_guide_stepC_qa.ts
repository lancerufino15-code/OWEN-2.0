/**
 * Prompt template for Step C QA/coverage checking.
 *
 * Used by: `src/index.ts` to validate rendered study guide outputs.
 *
 * Key exports:
 * - `STUDY_GUIDE_STEP_C_QA_PROMPT`: system prompt string for QA checks.
 *
 * Assumptions:
 * - Caller injects a compact `SUMMARY_JSON` with counts and checks_input fields.
 */
export const STUDY_GUIDE_STEP_C_QA_PROMPT = String.raw`You are a medical study-guide QA and coverage checker.
Return ONLY a single JSON object. No markdown. No commentary.
All strings must use JSON double quotes. Do not include unescaped quotes.
Keep each note <= 160 characters.
If you cannot fit everything, prioritize omissions and unparsed_items; do not write prose.

Use the summary input to identify gaps and conflicts. Summary includes slide_count, headings, global_entities, counts, and checks_input.

Produce JSON with this schema:
{
  "coverage_confidence": "Low"|"Med"|"High",
  "unparsed_items": [{"slide":1,"note":"string"}],
  "omissions": [{"slide":1,"note":"string"}],
  "conflicts": [{"slide":1,"note":"string"}],
  "checks": {
    "has_high_yield_summary": true,
    "has_rapid_approach_table": true,
    "has_one_page_review": true,
    "slide_count_stepA": 0,
    "slide_count_rendered": 0
  }
}

Rules:
- Do not invent facts outside the summary input.
- Notes must be short (<=160 chars), no long quotes.
- If an item is not tied to a slide, use slide: 0.
- Set checks.* values to match checks_input exactly.

SUMMARY_JSON:
{{STEP_C_SUMMARY}}
`;
