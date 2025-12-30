# src/prompts/study_guide_stepA_extract.txt

Purpose
- Legacy/plain-text prompt template for Step A extraction of slide facts.

Schema/fields
- lecture_title: string.
- chunk.start_slide/end_slide: slide range for this chunk.
- slides[n/page/sections/tables]: per-slide structured facts, tags, and tables.
- tags: constrained set of categories (disease, symptom, diagnostic, etc.).

Who reads/writes
- Humans (prompt design); not referenced by runtime code.

Notes
- NOTE: Current runtime uses the TypeScript prompt constant in `src/prompts/study_guide_stepA_extract.ts`.

Examples
- Placeholders: `{{CHUNK_START}}`, `{{CHUNK_END}}`.
