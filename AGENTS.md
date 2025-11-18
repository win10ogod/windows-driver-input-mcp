# Repository Guidelines

## Project Structure & Module Organization
- `main.py` — FastMCP server and tool definitions (Click/Type/State).
- `src/desktop/` — Windows automation core (PowerShell, UIA, input): `service.py`, `views.py`, `config.py`.
- `src/tree/` — UIA tree scanning + screenshot annotation: `service.py`, `views.py`, `utils.py`, `config.py`.
- `assets/` — logos, demos, screenshots (place PR screenshots in `assets/screenshots/`).
- `tests/` — `pytest` suites live here.
- `manifest.json`, `server.json` — MCP packaging and registry metadata.
- `IbInputSimulator/` — vendored reference; avoid edits unless necessary.

## Build, Test, and Development Commands
- Prereqs: Windows 7–11, Python `3.13+`, `uv` installed.
- Install deps: `uv sync`.
- Run (stdio): `uv run main.py --transport stdio`.
- Run (SSE): `uv run main.py --transport sse --host localhost --port 8000`.
- Tests: `uv run -m pytest`.
- Lint/format: `uvx ruff format . && uvx ruff check .`.
- Pack for Claude Desktop: `npx @anthropic-ai/mcpb pack`.

## Coding Style & Naming Conventions
- Python, 4‑space indent, PEP 8; type hints for public APIs.
- Names: modules/functions `snake_case`, classes `PascalCase`, constants `UPPER_SNAKE_CASE`.
- Docstrings: Google style (Args, Returns, Raises). Keep functions cohesive.
- Imports order: stdlib → third‑party → local.
- Use `logging`; never `print` in library code.

## Testing Guidelines
- Framework: `pytest`. Name files `test_*.py` in `tests/`.
- Mock UIA/`pyautogui`; prefer deterministic tests around `src/tree` utilities.
- Add regression tests for new tools (window states, localization).
- Run via `uv run -m pytest`; keep tests fast and isolated.

## Commit & Pull Request Guidelines
- Commits: clear, imperative messages (e.g., `feat:`, `fix:`).
- PRs include purpose/scope, linked issues, and local test steps.
- Include screenshots/GIFs for UX changes (save under `assets/screenshots/`).
- Note Windows version/localization impact and risk.
- Keep MCP tool names/signatures stable; if changed, update `manifest.json` and `README.md`.

## Security & Configuration Tips
- This server can control your desktop and run PowerShell—use on non‑critical machines/VMs.
- Prefer English UI for best UIA results; for other languages, consider disabling Launch/Switch tools.
- Avoid editing `IbInputSimulator/` unless required.

