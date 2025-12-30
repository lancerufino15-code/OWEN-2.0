/**
 * Floating model switcher widget for the chat UI (legacy/root asset copy).
 *
 * Used by: root `index.html` to allow users to select a model and persist it
 * in localStorage, while patching `/api/chat` requests with the selection.
 *
 * Key exports:
 * - None (self-invoking module that registers UI and fetch interception).
 *
 * Assumptions:
 * - Runs in a browser environment and can safely monkey-patch `window.fetch`.
 */
(() => {
  const STORE_KEY = "owen.model";
  let selected = localStorage.getItem(STORE_KEY) || "gpt-4-turbo";
  const legacyMap = {
    "gpt-4": "gpt-4o",
    "sora-2": "gpt-image-1",
    "gpt-5": "gpt-4-turbo",
  };
  if (legacyMap[selected]) {
    selected = legacyMap[selected];
    localStorage.setItem(STORE_KEY, selected);
  }

  // ---------- UI (Shadow DOM, floating pill) ----------
  const host = document.createElement("div");
  host.id = "owen-model-switcher-root";
  Object.assign(host.style, {
    position: "fixed",
    bottom: "16px",
    right: "16px",
    zIndex: "2147483647", // top
  });
  document.addEventListener("DOMContentLoaded", () => document.body.appendChild(host));
  const root = host.attachShadow({ mode: "open" });
  root.innerHTML = `
    <style>
      .box {
        font: 13px/1.4 var(--font-body, system-ui, -apple-system, Segoe UI, Roboto, sans-serif);
        color: var(--text-primary);
        background: var(--bg-surface-alt); border: 1px solid var(--border-subtle); border-radius: 999px;
        padding: 6px 10px; display: inline-flex; gap: 8px; align-items: center;
        box-shadow: var(--shadow-soft);
      }
      .label { color: var(--text-secondary); }
      select {
        border: 1px solid var(--border-subtle); border-radius: 999px; padding: 4px 8px;
        background: var(--bg-surface); color: var(--text-primary); outline: none;
      }
      button {
        border: 1px solid var(--border-subtle); border-radius: 999px; padding: 4px 8px;
        background: var(--bg-surface); color: var(--text-primary); cursor: pointer;
      }
      select:hover,
      button:hover {
        background: var(--bg-surface-hover);
      }
    </style>
    <div class="box" title="Choose ChatGPT model (does not change your page layout)">
      <span class="label">🧠 Model</span>
      <select id="owen-model"></select>
      <button id="owen-hide" aria-label="Hide switcher">✕</button>
    </div>
  `;

  const $ = (q) => root.querySelector(q);
  const sel = $("#owen-model");

  $("#owen-hide").addEventListener("click", () => host.remove());

  async function loadModels() {
    // Ask your Worker for allowed models; if not present, fall back safely.
    try {
      const r = await fetch("/api/models");
      if (!r.ok) throw new Error("no models api");
      const payload = await r.json();
      const models = Array.isArray(payload)
        ? payload
        : Array.isArray(payload?.models)
          ? payload.models
          : null;
      if (!models?.length) throw new Error("empty");
      return models.map(m => {
        if (typeof m === "string") {
          const isImage = ["gpt-image-1", "gpt-image-1-mini", "dall-e-3", "dall-e-2"].includes(m);
          const label = isImage ? `${m} (image)` : m;
          return { id: m, label };
        }
        return m;
      });
    } catch {
      // Fallback list (won’t break UI if /api/models is missing)
      return [
        { id: "gpt-4-turbo", label: "gpt-4-turbo" },
        { id: "gpt-4.1", label: "gpt-4.1" },
        { id: "gpt-4.1-mini", label: "gpt-4.1-mini" },
        { id: "gpt-5", label: "gpt-5" },
        { id: "gpt-5-mini", label: "gpt-5-mini" },
        { id: "gpt-4o", label: "gpt-4o" },
        { id: "gpt-image-1", label: "gpt-image-1 (image)" },
        { id: "gpt-image-1-mini", label: "gpt-image-1-mini (image)" },
        { id: "dall-e-3", label: "dall-e-3 (image)" },
        { id: "dall-e-2", label: "dall-e-2 (image)" },
      ];
    }
  }

  loadModels().then(models => {
    sel.innerHTML = "";
    for (const m of models) {
      const opt = document.createElement("option");
      opt.value = m.id; opt.textContent = m.label || m.id;
      sel.appendChild(opt);
    }
    // pick prior selection if it exists, else default to first available
    const found = models.some(m => m.id === selected);
    sel.value = found ? selected : (models[0]?.id || "gpt-4-turbo");
    selected = sel.value;
    localStorage.setItem(STORE_KEY, selected);
  });

  sel.addEventListener("change", () => {
    selected = sel.value;
    localStorage.setItem(STORE_KEY, selected);
  });

  // ---------- Fetch patch (surgical) ----------
  // Intercept POST /api/chat with JSON body and add { model: <selected> } if absent.
  // Does not touch other requests; leaves streaming/SSE handling to your existing code.
  const originalFetch = window.fetch;
  window.fetch = function(input, init) {
    return (async () => {
      try {
        const url = typeof input === "string" ? input : (input && input.url) || "";
        const method = (init?.method || (input && input.method) || "GET").toUpperCase();
        // Only patch our chat endpoint
        if (method === "POST" && url.includes("/api/chat")) {
          // Normalize headers
          const headers = new Headers(init?.headers || (input instanceof Request ? input.headers : undefined));
          const ct = (headers.get("content-type") || "").toLowerCase();

          // If JSON body, merge model; if not, add ?model=... to URL
          if (ct.includes("application/json") && init?.body && typeof init.body === "string") {
            try {
              const obj = JSON.parse(init.body);
              if (!obj.model) obj.model = selected;
              init = { ...init, headers, body: JSON.stringify(obj) };
            } catch {
              // Fall through to query param path
              input = appendModelParam(input, selected);
            }
          } else {
            input = appendModelParam(input, selected);
          }
        }
      } catch {
        // fail open—do nothing
      }
      return originalFetch(input, init);
    })();
  };

  function appendModelParam(input, model) {
    if (typeof input === "string") {
      const u = new URL(input, location.origin);
      if (!u.searchParams.get("model")) u.searchParams.set("model", model);
      return u.toString();
    }
    if (input instanceof Request) {
      const u = new URL(input.url);
      if (!u.searchParams.get("model")) u.searchParams.set("model", model);
      return new Request(u.toString(), input);
    }
    return input;
  }
})();
