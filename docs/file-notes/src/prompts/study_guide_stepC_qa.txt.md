# src/prompts/study_guide_stepC_qa.txt

Purpose
- Legacy/plain-text prompt template for Step C QA/coverage checks.

Schema/fields
- coverage_confidence: Low/Med/High.
- unparsed_items/omissions/conflicts: arrays of {slide, note}.
- checks: required booleans + slide counts.

Who reads/writes
- Humans (prompt design); not referenced by runtime code.

Notes
- NOTE: Current runtime uses the TypeScript prompt constant in `src/prompts/study_guide_stepC_qa.ts`.

Examples
- Placeholder: `{{STEP_C_SUMMARY}}`.
