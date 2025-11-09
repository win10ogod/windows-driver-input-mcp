# Windows Driver Input MCP

Standalone MCP server exposing driver-level keyboard/mouse tools via IbInputSimulator.

- No UIA tree, no vision, no shell — just input
- Backends: DLL first (`ibsim-dll`, no AHK required); optional AHK (`ibsim-ahk`, requires AutoHotkey v2 64-bit)
- Supports Unicode text, WASD, combos, holds, and scroll

Repository: https://github.com/win10ogod/windows-driver-input-mcp

License: MIT

## Run

- stdio: `uv run main.py --transport stdio`
- SSE: `uv run main.py --transport sse --host localhost --port 8001`

Configure via env (defaults shown):

- `WINDOWS_MCP_INPUT_BACKEND=ibsim-dll` (or `ibsim-ahk`)
- `WINDOWS_MCP_INPUT_DRIVER=AnyDriver`
- `WINDOWS_MCP_RATE_MOVE_HZ=120`, `WINDOWS_MCP_RATE_MAX_DELTA=60`, `WINDOWS_MCP_RATE_SMOOTH=0.0`
- `WINDOWS_MCP_RATE_CPS=8.0`, `WINDOWS_MCP_RATE_KPS=12.0`
- `WINDOWS_INPUT_LOG_LEVEL=INFO`
- `IBSIM_DIR` optionally to point to `IbInputSimulator` directory if not colocated

## Tools

- `Input-Info`
- `Move-Tool`
- `Click-Tool`
- `Drag-Tool`
- `Scroll-Tool`
- `Type-Tool(text, method=unicode|clipboard|vk, press_enter=false)`
- `Shortcut-Tool`
- `Key-Tool(mode=tap|down|up|hold, key, times, interval_ms, hold_ms)`
- `Combo-Tool(keys=[...], hold_ms=...)`
- `Input-RateLimiter-Config`

## IbInputSimulator

Place `IbInputSimulator/Binding.AHK2/IbInputSimulator.dll` and `IbInputSimulator.ahk` either:
- Inside this folder (`input_driver_server/IbInputSimulator/Binding.AHK2/...`), or
- Set `IBSIM_DIR` to the directory holding `Binding.AHK2` (or its parent).

The server auto-discovers common locations and also checks the parent repo layout, but for a truly standalone package, vendoring `IbInputSimulator` in this folder is recommended.

## Install

- Requirements
  - Windows 7–11
  - Python 3.13+ (`python --version`)
  - UV package manager: `pip install uv` (or see https://github.com/astral-sh/uv)
  - IbInputSimulator is already vendored in `input_driver_server/IbInputSimulator`.
  - Optional (only when using `ibsim-ahk`): AutoHotkey v2 64‑bit

- Setup (clone once)
  - `git clone https://github.com/win10ogod/windows-driver-input-mcp.git`
  - `cd windows-driver-input-mcp`
  - (Optional) `uv sync` if you prefer a local venv with pinned deps

## Client Config Snippets (JSON)

Use these JSON blocks in clients that accept Model Context Protocol server configs.

### Claude Desktop (Advanced > Add Local MCP)

Paste this JSON into the “Install from JSON” dialog:

```json
{
  "command": "uv",
  "args": [
    "--directory",
    "<ABSOLUTE PATH TO>/windows-driver-input-mcp",
    "run",
    "main.py"
  ],
  "env": {
    "WINDOWS_MCP_INPUT_BACKEND": "ibsim-dll",
    "WINDOWS_MCP_INPUT_DRIVER": "AnyDriver",
    "WINDOWS_MCP_RATE_MOVE_HZ": "120",
    "WINDOWS_MCP_RATE_MAX_DELTA": "60",
    "WINDOWS_MCP_RATE_SMOOTH": "0.0",
    "WINDOWS_MCP_RATE_CPS": "8.0",
    "WINDOWS_MCP_RATE_KPS": "12.0",
    "WINDOWS_INPUT_LOG_LEVEL": "INFO"
  }
}
```

Or install as an Extension package:

1) `cd windows-driver-input-mcp`

2) `npx @anthropic-ai/mcpb pack`

3) In Claude Desktop: Settings → Extensions → Advanced Settings → Install Extension → pick the generated `.mcpb`.

### Codex CLI (TOML)

Edit `%USERPROFILE%/.codex/config.toml` and add:

```toml
[mcp_servers.windows-driver-input]
command = "uv"
args = [
  "--directory",
  "<ABSOLUTE PATH TO>/windows-driver-input-mcp",
  "run",
  "main.py"
]
env = {
  WINDOWS_MCP_INPUT_BACKEND = "ibsim-dll",
  WINDOWS_MCP_INPUT_DRIVER = "AnyDriver",
  WINDOWS_MCP_RATE_MOVE_HZ = "120",
  WINDOWS_MCP_RATE_MAX_DELTA = "60",
  WINDOWS_MCP_RATE_SMOOTH = "0.0",
  WINDOWS_MCP_RATE_CPS = "8.0",
  WINDOWS_MCP_RATE_KPS = "12.0",
  WINDOWS_INPUT_LOG_LEVEL = "INFO"
}
```

Save and restart Codex CLI to load the new server.

### Gemini CLI (JSON)

Edit `%USERPROFILE%/.gemini/settings.json` and merge the following under `mcpServers`:

```json
{
  "mcpServers": {
    "windows-driver-input": {
      "command": "uv",
      "args": [
        "--directory",
        "<ABSOLUTE PATH TO>/windows-driver-input-mcp",
        "run",
        "main.py"
      ],
      "env": {
        "WINDOWS_MCP_INPUT_BACKEND": "ibsim-dll",
        "WINDOWS_MCP_INPUT_DRIVER": "AnyDriver",
        "WINDOWS_MCP_RATE_MOVE_HZ": "120",
        "WINDOWS_MCP_RATE_MAX_DELTA": "60",
        "WINDOWS_MCP_RATE_SMOOTH": "0.0",
        "WINDOWS_MCP_RATE_CPS": "8.0",
        "WINDOWS_MCP_RATE_KPS": "12.0",
        "WINDOWS_INPUT_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

Restart Gemini CLI to apply the configuration.

## Acknowledgements

- Portions of functionality are adapted from Windows-MCP:
  https://github.com/CursorTouch/Windows-MCP
- Special thanks to IbInputSimulator for driver-level input:
  https://github.com/Chaoses-Ib/IbInputSimulator
