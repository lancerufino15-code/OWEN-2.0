# tsconfig.json

Purpose
- TypeScript compiler configuration for the Worker source (`src/`).

Key fields
- compilerOptions.target/lib: ES2023 + WebWorker runtime targets.
- module/moduleResolution: ESNext with bundler resolution (Wrangler-compatible).
- strict/noUncheckedIndexedAccess: stricter type safety.
- types: includes `@cloudflare/workers-types` for Worker globals.
- resolveJsonModule/isolatedModules/noEmit: allow JSON imports, per-file builds, and no output.
- include: `src` only.
- exclude: `worker-configuration.d.ts` (generated type file).

Who reads/writes
- TypeScript compiler and editor tooling.
- Developers update when changing build targets or runtime assumptions.

Examples
- `tsc --noEmit` uses these settings via `npm run check`.
