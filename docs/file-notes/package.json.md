# package.json

Purpose
- Node package manifest for the OWEN Worker project (dependencies, scripts, template metadata).

Key fields
- name/description/private: package identity and privacy flag.
- cloudflare: template metadata for Cloudflare (label, products, categories, docs URLs, preview assets).
- dependencies: runtime deps bundled into the Worker (e.g., `pdfjs-dist`).
- devDependencies: tooling for tests/typegen/dev (wrangler, vitest, typescript, jsdom).
- scripts:
  - cf-typegen: generate Worker type defs (`wrangler types`).
  - dev/start: run local worker (`wrangler dev`).
  - deploy: deploy to Cloudflare (`wrangler deploy`).
  - check: typecheck + dry-run deploy.
  - test: run Node tests with the TS loader.

Who reads/writes
- npm/yarn/bun read dependencies and scripts.
- Wrangler uses scripts and cloudflare metadata.
- Humans update versions and scripts.

Examples
- `npm run dev`
- `npm run test`
