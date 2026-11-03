# Contributing

- **Definition of done.** A feature counts as done when it works, is reachable from the UI, has automated tests, and its limits are written down. A button that exists in the UI does not count. [IMPLEMENTATION_STATUS.md](IMPLEMENTATION_STATUS.md) only lists what has a passing test.
- **Before a commit.** Run `npm run typecheck && npm test && npm run e2e:headless`. The E2E suite also fails on any Content-Security-Policy violation or uncaught renderer error.
- **The core stays pure.** Code in `src/core` has no DOM and no Electron. It must run in the worker and under vitest.
- **New worksheet functions.** Add them to the matching `src/core/engine/functions/<category>.ts` with cases in `tests/unit/functions/`. Expected values come from Microsoft's documentation examples; note the source in the test. Regenerate the catalog with `node scripts/gen-catalog.mjs` when Microsoft adds functions.
- **No invented results.** Never return a placeholder result. If something can't be calculated, keep the file's value and report it, or return the error Excel would.
- **Shortcuts.** Add new shortcuts to `src/renderer/app/shortcuts.ts`, which drives dispatch, the shortcut dialog and the reference document.
