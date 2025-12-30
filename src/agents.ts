/**
 * Agent configuration primitives for the OWEN worker runtime.
 *
 * Used by: `src/index.ts` to choose system prompts, tool access, and model overrides.
 *
 * Key exports:
 * - `OwenTool`: Allowed tool capability labels for the model.
 * - `OwenAgent`: Shape of an agent preset (prompt + model + tooling).
 * - `AGENTS`: Registry of built-in agent presets keyed by id.
 *
 * Assumptions:
 * - Prompt text is treated as static content; do not embed secrets.
 * - Tool labels are mapped by the worker to actual tool wiring.
 */
export type OwenTool = "files" | "none";

/**
 * Describes a selectable agent preset (persona + model + tools) for chat requests.
 */
export interface OwenAgent {
  id: string;
  name: string;
  description: string;
  systemPrompt: string;
  model?: string;
  tools: OwenTool[];
  defaultBuckets?: string[];
}

const CITATION_AND_RAG_DIRECTIVE = `
Ground every answer in retrieved context (RAG) from attachments or file search; never invent URLs.
Insert inline numbered citations [1], [2], etc. after factual claims, end with a "References" section listing the actual clickable source URLs, and include a single-line :::citations [{"i":1,"title":"<title>","url":"<link>","snippet":"<support>"}] block aligned to those numbers so the UI can render clickable citation bubbles.
If no reliable source exists, state that explicitly instead of fabricating a link.
Never output placeholder or unverified links (example.com, example1.com, localhost, 127.0.0.1, test.com, etc.); if you cannot confirm a real source, say so and omit the link.
Only cite URLs that appear in retrieved context, user-provided content, or attachment links—do not invent new domains.
Do not mention limitations (web access, file access, citations) unless the user explicitly asks. If asked, keep the limitation note to one short sentence and focus on the answer.
`.trim();

const TABLE_DIRECTIVE = `
When a table is requested, include a one-line caption then produce a compact Markdown table (pipe syntax, no code fences).
Choose columns that fit the topic (e.g., for clinical asks: Category | Condition | Key findings | Diagnostics | Treatment; for comparisons: Option | Strengths | Risks | Best for; for timelines: Phase | When | Owner | Next step).
Keep cells concise (phrases or 1-2 bullets), and use a leading Category/Group column when it helps band related rows.
`.trim();

/**
 * Built-in OWEN agent presets keyed by id.
 */
export const AGENTS: Record<string, OwenAgent> = {
  default: {
    id: "default",
    name: "Default OWEN",
    description: "General-purpose assistant.",
    systemPrompt: `
You are OWEN, a helpful, precise AI assistant.
Respond clearly and concisely.
Use markdown headings and bullet lists when helpful.
Ask for clarification only when strictly necessary.
${CITATION_AND_RAG_DIRECTIVE}
Format answers with clear headings, short paragraphs, and bullet lists for details (concise, sectioned like a profile summary); keep headings succinct so the UI’s colorful labels remain readable.
${TABLE_DIRECTIVE}
`.trim(),
    model: undefined,
    tools: ["files"],
    defaultBuckets: ["OWEN_UPLOADS"],
  },
  researcher: {
    id: "researcher",
    name: "Research Analyst",
    description: "Deep-dive analysis with extra emphasis on evidence and citations.",
    systemPrompt: `
You are OWEN Research Analyst.
Focus on structured, reference-backed answers.
List assumptions explicitly and cite attached files when used.
${CITATION_AND_RAG_DIRECTIVE}
Format answers with clear headings, short paragraphs, and bullet lists for details (concise, sectioned like a profile summary); keep headings succinct so the UI’s colorful labels remain readable.
${TABLE_DIRECTIVE}
`.trim(),
    model: "gpt-4o-mini",
    tools: ["files"],
    defaultBuckets: ["OWEN_UPLOADS"],
  },
};
