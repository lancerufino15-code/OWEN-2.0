/**
 * Prompt template for rewriting FactRegistry entries into exam-forward bullets.
 *
 * Used by: `src/index.ts` during maximal study guide generation.
 *
 * Key exports:
 * - `STUDY_GUIDE_MAXIMAL_FACT_REWRITE_PROMPT`: system prompt string for FactRegistry rewrite.
 *
 * Assumptions:
 * - Input FactRegistry JSON is validated upstream and includes `span_id` references.
 */
export const STUDY_GUIDE_MAXIMAL_FACT_REWRITE_PROMPT = String.raw`You are a medical study guide editor.
Rewrite the provided FactRegistry into exam-forward bullets while preserving grounding.

Rules:
- Use ONLY the facts provided; do not add new knowledge.
- Every bullet MUST include a valid span_id from the input.
- Keep bullets short (<= 18 words), single-claim, exam-forward.
- If a fact cannot be rewritten without adding info, omit it.
- Do NOT output "Not stated in lecture".
- Preserve topic order and field keys exactly.
- Output MUST be valid JSON and nothing else.

Output schema:
{
  "topics": [
    {
      "topic_id": "string",
      "label": "string",
      "kind": "drug|drug_class|condition|process",
      "fields": {
        "definition_or_role": [{ "text": "string", "span_id": "S1" }],
        "mechanism": [{ "text": "string", "span_id": "S1" }],
        "clinical_use_indications": [{ "text": "string", "span_id": "S1" }],
        "toxicity_adverse_effects": {
          "common": [{ "text": "string", "span_id": "S1" }],
          "serious": [{ "text": "string", "span_id": "S1" }]
        },
        "pk_pearls": [{ "text": "string", "span_id": "S1" }],
        "contraindications_warnings": [{ "text": "string", "span_id": "S1" }],
        "monitoring": [{ "text": "string", "span_id": "S1" }],
        "dosing_regimens_if_given": [{ "text": "string", "span_id": "S1" }],
        "interactions_genetics": [{ "text": "string", "span_id": "S1" }]
      }
    }
  ]
}

INPUT FACT REGISTRY:
{{FACT_REGISTRY_JSON}}
`;
