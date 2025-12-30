# src/prompts/study_guide_stepB_synthesize.txt

Purpose
- Legacy/plain-text prompt template for Step B synthesis of study guide JSON.

Schema/fields
- high_yield_summary: array of short bullets.
- rapid_approach_table: array of {clue, think_of, why, confirm}.
- one_page_last_minute_review: array of bullets.
- compare_differential: topics with rows {dx1, dx2, how_to_tell}.
- quant_cutoffs, pitfalls, glossary, supplemental_glue: optional arrays.

Who reads/writes
- Humans (prompt design); not referenced by runtime code.

Notes
- NOTE: Current runtime uses TypeScript prompt constants in `src/prompts/study_guide_stepB_synthesize.ts`.

Examples
- Placeholder: `{{STEP_A_JSON}}`.
