/**
 * Client-side helper for parsing OpenAI/Workers AI SSE streams.
 *
 * Used by: `src/index.ts` for local dev flows and any browser-side streaming
 * consumers that need to parse `data:` SSE frames into text tokens.
 *
 * Key exports:
 * - `streamChat`: Performs POST + SSE parse and invokes token callbacks.
 *
 * Assumptions:
 * - Server responds with `text/event-stream` framing and `[DONE]` or `__END__`.
 * - Payloads may be either legacy chat-completions or Responses API shapes.
 */
type AnyObj = Record<string, any>;

function extractText(obj: AnyObj, eventName?: string): string | null {
  if (typeof obj?.delta === "string") return obj.delta;
  if (typeof obj?.text === "string") return obj.text;

  const d = obj?.delta;
  if (d?.content && Array.isArray(d.content)) {
    for (const part of d.content) {
      if ((part.type === "text" || part.type === "output_text") && typeof part.text === "string") {
        return part.text;
      }
    }
  }

  if (Array.isArray(obj?.choices)) {
    const chunk = obj.choices[0]?.delta?.content;
    if (typeof chunk === "string") return chunk;
  }

  if (eventName === "response.completed" || obj?.type === "response.completed") {
    if (typeof obj?.output_text === "string") return obj.output_text;
    const parts = obj?.response?.output?.[0]?.content;
    if (Array.isArray(parts)) {
      const all = parts.map((p: any) => p.text).filter(Boolean).join("");
      if (all) return all;
    }
  }
  return null;
}

/**
 * Post a request to an SSE endpoint and stream parsed text tokens.
 *
 * @param endpoint - URL that accepts JSON payloads and streams SSE events.
 * @param payload - Request body to POST.
 * @param onToken - Callback for each parsed text token chunk.
 * @param onDone - Optional callback when a done marker is received.
 * @param onError - Optional callback for parse/network errors.
 * @param signal - Optional AbortSignal to cancel the fetch.
 * @returns Resolves when the stream completes or rejects on errors.
 * @remarks Side effects: performs a network fetch and consumes a ReadableStream.
 */
export async function streamChat({
  endpoint,
  payload,
  onToken,
  onDone,
  onError,
  signal,
}: {
  endpoint: string;
  payload: any;
  onToken: (t: string) => void;
  onDone?: () => void;
  onError?: (e: any) => void;
  signal?: AbortSignal;
}) {
  const res = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    signal,
  });

  if (!res.ok || !res.body) {
    const txt = await res.text().catch(() => "");
    throw new Error(`HTTP ${res.status} ${res.statusText} ${txt}`);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buf = "";
  let gotAnyText = false;
  let lastCompletedText = "";

  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      const events = buf.split(/\r?\n\r?\n/);
      buf = events.pop() || "";

      for (const evt of events) {
        let eventName: string | undefined;
        let dataPayload = "";

        for (const line of evt.split(/\r?\n/)) {
          if (!line) continue;
          const idx = line.indexOf(":");
          const field = idx === -1 ? line.trim() : line.slice(0, idx).trim();
          const valuePart = idx === -1 ? "" : line.slice(idx + 1).trim();
          if (field === "event") eventName = valuePart;
          if (field === "data") dataPayload += (dataPayload ? "\n" : "") + valuePart;
        }

        if (!dataPayload) continue;
        if (dataPayload === "[DONE]" || dataPayload === "__END__") {
          onDone?.();
          return;
        }

        let obj: AnyObj | string = dataPayload;
        try {
          obj = JSON.parse(dataPayload);
        } catch {
          onToken(dataPayload);
          gotAnyText = true;
          continue;
        }

        if (typeof obj === "object") {
          const maybeCompleted =
            (obj as AnyObj).output_text ??
            (obj as AnyObj)?.response?.output?.[0]?.content?.map((p: any) => p.text).join("");
          if (maybeCompleted) lastCompletedText = String(maybeCompleted);
        }

        const t = extractText(obj as AnyObj, eventName);
        if (t) {
          onToken(t);
          gotAnyText = true;
        }
      }
    }

    if (!gotAnyText && lastCompletedText) {
      onToken(lastCompletedText);
      gotAnyText = true;
    }

    if (!gotAnyText) {
      throw new Error("Parser received SSE but extracted no text (check event names & shapes).");
    }

    onDone?.();
  } catch (e) {
    onError?.(e);
    throw e;
  } finally {
    reader.releaseLock?.();
  }
}
