# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# GurobiStudio

Excel for Mac add-in (Office.js) plus a local Python helper that runs gurobipy models.

## Commands

- Install or update everything: `scripts/install.sh` (env: `GS_PORT`, `GS_PLACEMENT=data|custom|home`, `GS_LAUNCH_AGENT=1`)
- Run the helper: `scripts/run-helper.sh`. It doesn't auto-reload, so restart it after changing server code.
- Helper tests: `cd helper && .venv/bin/pytest`, one file: `.venv/bin/pytest tests/test_wire.py`, lint: `.venv/bin/ruff check .`
- Add-in: `cd add-in && npm run build`, `npm test`, one file: `npx vitest run src/core/sse.test.ts`, `npm run lint`
- Browser acceptance, real helper and Gurobi on a spare port: `cd add-in && npm run build && npm run e2e`
- Big-data acceptance, 1M cells end to end: `cd add-in && npm run build && npm run e2e:big` (`GS_BIG_SCALE=0.01` for a quick run)
- Open the pane in a browser: `https://localhost:8765/?k=<page key>`, with the key from the `page_key` file in the data folder. The helper serves `add-in/dist`, so build first.
- Regenerate ribbon icons: `python3 scripts/make_icons.py add-in/public/icons`
- After manifest changes: re-run install, quit Excel, clear the cache (see docs/TESTING.md)

## Architecture

The helper is a FastAPI app on `127.0.0.1` over HTTPS. It serves both the built add-in and the API. Its data folder is `~/Library/Application Support/GurobiStudio/` (override with `GS_DATA_DIR`) and holds `token`, `page_key`, `port`, `settings.json`, `trusted_models.json` and `runs/`.

A solve crosses both sides:

1. `add-in/src/core/solve.ts` reads cells through the adapter and sends `POST /api/run`. A 409 means either `trust_required` or `busy`.
2. `RunManager` (`helper/gurobistudio/runs.py`) starts `python -m gurobistudio.runner RUN_DIR` in its own process group. Only one run goes at a time, and the last 10 run folders are kept.
3. `runner.py` decodes the items (`wire.py`), exposes them as model globals and through the `SolverStudio` shim, then `exec`s the model. It compares the items to a deep-copied snapshot and writes `result.json`. Results never go through stdout, so a model's `print` can't corrupt them.
4. `RunManager` turns stdout into SSE log events, dropping lines that match `CREDENTIAL_LINE`, then adds a final `result`, `error` or `stopped` event.
5. `solve.ts` parses the stream with `fetch` (`EventSource` can't send the auth header) and writes changed items back to the sheet.

Stop sends SIGINT to the process group so Gurobi releases its license token, then SIGKILL after 3 seconds.

Security layers:

- `SecurityMiddleware` checks Host, Origin and the bearer token.
- The token is injected into the served page only when the request has the correct `?k=` page key.
- Model approvals live in the data folder, never the workbook, so a workbook can't approve itself.

In `add-in/src/`, `excel/` holds the `ExcelAdapter` interface with its Office and mock implementations. `main.tsx` picks one at startup; outside Excel it uses the mock and exposes `window.mockCells` for browser tests. `core/` holds non-UI logic and `ui/` holds React with a small store. `core/persist.ts` saves models, items and settings to a custom XML part in the workbook. `examples/*.json` are bundled into the add-in. `commands.ts` handles ribbon actions in the shared runtime.

## Current state

- Milestones M0 to M6 are built on branch `m4-m6`.
- `docs/TESTING.md` lists manual Excel checks and which have been run in real Excel.
- `docs/DECISIONS.md` refers to "the brief", which isn't in this repo.

## Rules

- Never print, log or commit the WLS secret. Read `gurobi.lic` only to check for leaks, without printing it.
- All Office.js calls go through `add-in/src/excel/` (from M3), so the UI runs on the mock adapter. Current exceptions: `main.tsx` and `commands.ts` (startup and ribbon wiring) and `core/diagnostics.ts`.
- Shape building lives in TypeScript; decoding and change detection live in Python.
- Pinned versions only. TypeScript stays below 6.1 until typescript-eslint supports newer.
- Record every fallback and open choice in `docs/DECISIONS.md`.
- Browser automation uses the Playwright CLI, with scripts outside the repo unless kept on purpose.
