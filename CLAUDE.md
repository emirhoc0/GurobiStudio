# GurobiStudio

Excel for Mac add-in (Office.js) plus a local Python helper that runs gurobipy models.

## Commands

- Install or update everything: `scripts/install.sh` (env: `GS_PORT`, `GS_PLACEMENT=data|custom|home`)
- Run the helper: `scripts/run-helper.sh`
- Helper tests: `cd helper && .venv/bin/pytest`, lint: `.venv/bin/ruff check .`
- Add-in: `cd add-in && npm run build`, `npm test`, `npm run lint`
- Regenerate ribbon icons: `python3 scripts/make_icons.py add-in/public/icons`
- After manifest changes: re-run install, quit Excel, clear the cache (see docs/TESTING.md)

## Rules

- Never print, log or commit the WLS secret. Read `gurobi.lic` only to check for leaks, without printing it.
- All Office.js calls go through `add-in/src/excel/` (from M3), so the UI runs on the mock adapter.
- Shape building lives in TypeScript; decoding and change detection live in Python.
- Pinned versions only. TypeScript stays below 6.1 until typescript-eslint supports newer.
- Record every fallback and open choice in `docs/DECISIONS.md`.
- Browser automation uses the Playwright CLI, with scripts outside the repo unless kept on purpose.
