/**
 * Prompt builder for lecture quiz generation.
 *
 * Used by: `src/index.ts` when serving `/api/lecture/quiz` style endpoints.
 *
 * Key exports:
 * - `buildLectureQuizPrompt`: produces system + user prompt strings.
 *
 * Assumptions:
 * - Output must be strict JSON matching the schema enforced by the backend.
 */
type LectureQuizPromptOpts = {
  lectureId: string;
  lectureTitle?: string;
  context: string;
  referenceLabels?: string[];
};

const QUIZ_SYSTEM_PROMPT = [
  "You are OWEN generating a lecture quiz.",
  "Respond with a JSON object that matches the provided response schema.",
  "Do not output markdown, code fences, or commentary.",
].join("\n");

/**
 * Build system and user prompt strings for quiz generation.
 *
 * @param opts - Lecture metadata, context excerpts, and optional labels.
 * @returns Object containing system + user prompt strings.
 */
export function buildLectureQuizPrompt(opts: LectureQuizPromptOpts): { system: string; user: string } {
  const labels = (opts.referenceLabels || []).map(label => label.trim()).filter(Boolean);
  const referenceList = labels.length ? labels.map(label => `- ${label}`).join("\n") : "- (none)";
  const lectureTitle = (opts.lectureTitle || "").trim();

  const user = [
    "Task: Generate a lecture quiz using ONLY the excerpts below. Do not add outside facts.",
    "Constraints:",
    "- Generate exactly 5 questions (setSize = 5).",
    "- Use Amboss-style multiple choice with 4 or 5 options labeled A-E.",
    "- Include plausible distractors and exactly one best answer per question.",
    "- Provide a brief rationale grounded in the lecture text.",
    "- Include 1-3 short tags and a difficulty (easy|medium|hard) per question.",
    "- If a reference label fits, include 1-2 labels in references; otherwise use an empty array.",
    "",
    `Lecture ID: ${opts.lectureId}`,
    lectureTitle ? `Lecture Title: ${lectureTitle}` : "",
    "Reference labels:",
    referenceList,
    "",
    "Lecture excerpts:",
    opts.context || "",
  ]
    .filter(Boolean)
    .join("\n");

  return { system: QUIZ_SYSTEM_PROMPT, user };
}
