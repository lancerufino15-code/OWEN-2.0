# LLM Chat Application Template

A simple, ready-to-deploy chat application template powered by Cloudflare Workers AI. This template provides a clean starting point for building AI chat applications with streaming responses.

[![Deploy to Cloudflare](https://deploy.workers.cloudflare.com/button)](https://deploy.workers.cloudflare.com/?url=https://github.com/cloudflare/templates/tree/main/llm-chat-app-template)

<!-- dash-content-start -->

## Demo

This template demonstrates how to build an AI-powered chat interface using Cloudflare Workers AI with streaming responses. It features:

- Real-time streaming of AI responses using Server-Sent Events (SSE)
- Easy customization of models and system prompts
- Support for AI Gateway integration
- Clean, responsive UI that works on mobile and desktop

## Features

- 💬 Simple and responsive chat interface
- ⚡ Server-Sent Events (SSE) for streaming responses
- 🧠 Powered by Cloudflare Workers AI LLMs
- 🛠️ Built with TypeScript and Cloudflare Workers
- 📱 Mobile-friendly design
- 🔄 Maintains chat history on the client
- 🔎 Built-in Observability logging
<!-- dash-content-end -->

## Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) (v18 or newer)
- [Wrangler CLI](https://developers.cloudflare.com/workers/wrangler/install-and-update/)
- A Cloudflare account with Workers AI access

### Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/cloudflare/templates.git
   cd templates/llm-chat-app
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Generate Worker type definitions:
   ```bash
   npm run cf-typegen
   ```

### Development

Start a local development server:

```bash
npm run dev
```

This will start a local server at http://localhost:8787.

Note: Using Workers AI accesses your Cloudflare account even during local development, which will incur usage charges.

### Deployment

Deploy to Cloudflare Workers:

```bash
npm run deploy
```

### Monitor

View real-time logs associated with any deployed Worker:

```bash
npm wrangler tail
```

### Anki Upload Verification

```bash
curl -i \
  -F 'slidesPdf=@./slides.pdf;type=application/pdf' \
  -F 'transcriptTxt=@./transcript.txt;type=text/plain' \
  -F 'lectureTitle=Test' \
  -F 'lectureId=test_01' \
  https://<host>/api/anki/generate
```

## Project Structure

```
/
├── public/             # Static assets
│   ├── index.html      # Chat UI HTML
│   └── chat.js         # Chat UI frontend script
├── src/
│   ├── index.ts        # Main Worker entry point
│   └── types.ts        # TypeScript type definitions
├── test/               # Test files
├── wrangler.jsonc      # Cloudflare Worker configuration
├── tsconfig.json       # TypeScript configuration
└── README.md           # This documentation
```

## How It Works

### Backend

The backend is built with Cloudflare Workers and uses the Workers AI platform to generate responses. The main components are:

1. **API Endpoint** (`/api/chat`): Accepts POST requests with chat messages and streams responses
2. **Streaming**: Uses Server-Sent Events (SSE) for real-time streaming of AI responses
3. **Workers AI Binding**: Connects to Cloudflare's AI service via the Workers AI binding

### Frontend

The frontend is a simple HTML/CSS/JavaScript application that:

1. Presents a chat interface
2. Sends user messages to the API
3. Processes streaming responses in real-time
4. Maintains chat history on the client side

### Message Persistence & Citations

- Conversation history is stored in `localStorage` (`owen.conversation.<id>`); there is no server-side conversation DB in this template.
- Assistant messages persist rendering metadata so citations survive reloads:
  - `text`: the rendered markdown shown in the UI.
  - `renderedMarkdown`: optional override if the stored text differs from the rendered markdown.
  - `sources[]`: normalized list used for the Sources drawer (`id`, `url`, `title`, `domain`, `snippet`).
  - `citations[]`: mapping used to rebuild inline citation pills (same shape as `sources[]`).
  - `answerSegments[]`: when Responses API returns structured segments, the UI rehydrates from segments to keep citation placement deterministic.
- If stored metadata is missing (older conversations), citations remain visible but are treated as unverified/disabled.

## Citation Contract (Free-response)

- Free-response mode (no system-primed context and no attachments) always uses the Responses API with `web_search`.
- The backend returns structured `answerSegments` + `sources`; the UI renders citation pills from segments and a Sources list from `sources` only.
- URLs are taken only from web_search tool output. `MIN_DISTINCT_SOURCES` (or `FREE_RESPONSE_MIN_UNIQUE_SOURCES`, default 8) is a soft target; answers still return with a `warnings` entry when fewer sources are found. Set `ENFORCE_MIN_DISTINCT_SOURCES=1` to hard-enforce the minimum for specialized workflows.

## Customization

### Changing the Model

To use a different AI model, update the `MODEL_ID` constant in `src/index.ts`. You can find available models in the [Cloudflare Workers AI documentation](https://developers.cloudflare.com/workers-ai/models/).

### Using AI Gateway

The template includes commented code for AI Gateway integration, which provides additional capabilities like rate limiting, caching, and analytics.

To enable AI Gateway:

1. [Create an AI Gateway](https://dash.cloudflare.com/?to=/:account/ai/ai-gateway) in your Cloudflare dashboard
2. Uncomment the gateway configuration in `src/index.ts`
3. Replace `YOUR_GATEWAY_ID` with your actual AI Gateway ID
4. Configure other gateway options as needed:
   - `skipCache`: Set to `true` to bypass gateway caching
   - `cacheTtl`: Set the cache time-to-live in seconds

Learn more about [AI Gateway](https://developers.cloudflare.com/ai-gateway/).

### Modifying the System Prompt

The default system prompt can be changed by updating the `SYSTEM_PROMPT` constant in `src/index.ts`.

### Styling

The UI styling is contained in the `<style>` section of `public/index.html`. You can modify the CSS variables at the top to quickly change the color scheme.

## Resources

- [Cloudflare Workers Documentation](https://developers.cloudflare.com/workers/)
- [Cloudflare Workers AI Documentation](https://developers.cloudflare.com/workers-ai/)
- [Workers AI Models](https://developers.cloudflare.com/workers-ai/models/)

## Study Guide Diagnostics

- Enable debug payloads: set `OWEN_DEBUG=1` in your Worker environment.
- Diagnostics are persisted to KV (prefers `OWEN_DIAG_KV`, falls back to `DOCS_KV`) under `diagnostics/{requestId}.json`.
- Local repro: run `npm run dev`, call `/api/machine/study-guide` with `mode=maximal`, then inspect the KV key for the emitted `requestId`.

## Study Guide Developer Notes

How to add a new lecture-type mapping (drug lecture vs disease lecture):
1. Update topic classification in `src/study_guides/inventory.ts` to detect the new entities and filter headings.
2. Map lecture facts into the FactRegistry fields in `src/study_guides/fact_registry.ts` (adjust hint rules + MH7 filters).
3. Extend `src/study_guides/render_maximal_html.ts` to render MH7 tables and highlight classes for the new type.
4. Add or update tests in `study_guide_mh7_quality.test.ts` and `study_guide_maximal_mh7_integration.test.ts`.
