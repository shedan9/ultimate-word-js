# Repository Guidelines

## Project Structure & Module Organization

This pnpm workspace implements a high-fidelity Word/OOXML preview and editing engine for Chinese documents.

- `packages/*/src/`: `core` (units/utilities), `ooxml` (ZIP/XML), `model` (document/styles), `fonts` (metrics), `layout` (pagination), `render-dom` (rendering), and `view` (viewport coordinates/read-only interaction).
- `apps/playground/`: Vite debugging application.
- `apps/fidelity/`: Word truth-generation tools; `fixtures/` contains DOCX and coordinate truth JSON, with fixture specifications in `fixtures/src/`.
- `packages/fonts/packs/`: committed font metrics; proprietary font sources stay untracked.
- `docs/`: architecture, API design, and development plan. Consult `CLAUDE.md` for detailed implementation guidance.

## Build, Test, and Development Commands

Use Node from `.node-version` (24.19.0) and pnpm from `packageManager` (11.21.0). With fnm, expose the matching Node installation on `PATH` in noninteractive shells.

- `pnpm install --frozen-lockfile`: install locked dependencies.
- `pnpm --filter @uw/playground dev`: start the playground.
- `pnpm build`: build workspace packages and application through Turbo.
- `pnpm turbo run typecheck test`: run TypeScript checks and Vitest suites, matching CI.
- `pnpm --filter @uw/layout test`: run one package’s tests.
- `pnpm lint` / `pnpm lint:fix`: check / fix Biome issues.
- `pnpm format`: format supported files.
- `pnpm truth`: regenerate reference coordinates; requires Windows with Word COM access.

Workspace exports target source TypeScript, so development and tests need no preliminary build.

## Coding Style & Naming Conventions

Use strict TypeScript ESM, explicit `import type`, and `.ts` extensions on relative imports. Avoid enums and parameter properties. Biome enforces two-space indentation, single quotes, semicolons, trailing commas, and 110-character lines. Use kebab-case filenames, camelCase functions/variables, and PascalCase types. Write Chinese comments explaining reasoning.

## Architecture Constraints

Keep package dependencies directional and stage outputs structured-cloneable. Keep DOM APIs out of `layout`; inject `TextMeasurer`. Compute layout in twips through core unit helpers; convert at rendering boundaries. Preserve unknown XML for round trips.

## Testing Guidelines

Colocate Vitest tests as `src/*.test.ts`. Add focused regressions for behavior changes. No percentage coverage threshold is configured. Preserve coordinate tolerances and never lower `MIN_L2_MATCH`. Commit fixture DOCX and truth JSON together; exclude generated PDFs and `out/` artifacts.

## Commit & Pull Request Guidelines

Follow history’s `type(scope): explanation` pattern, e.g. `fix(layout,model): ...`; explain why. PRs should describe behavior changes, link relevant issues, and report validation. Include truth comparisons for layout changes and screenshots for visible changes. Update affected architecture/API documentation and development progress.
