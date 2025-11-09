import os
import logging
import click
import ctypes
from ctypes import wintypes
from contextlib import asynccontextmanager
from fastmcp import FastMCP
from typing import Literal

# Local imports
from backend import (
    IBSimulatorDLLBackend,
    IBSimulatorAHKBackend,
    InputBackend,
)
from rate import RateLimiter, RateConfig


logger = logging.getLogger(__name__)
logging.basicConfig(level=os.getenv("WINDOWS_INPUT_LOG_LEVEL", "INFO").upper())


def _coerce_xy(value) -> tuple[int, int]:
    """Coerce various loc-like inputs to a pair of ints (x, y).

    Accepts:
    - [x,y] list/tuple
    - {"x":..., "y":...} dict
    - string forms like "800,600", "800 600", "[800,600]"
    Raises ValueError for invalid inputs.
    """
    # List/Tuple
    if isinstance(value, (list, tuple)) and len(value) == 2:
        return int(value[0]), int(value[1])
    # Dict with x/y
    if isinstance(value, dict) and "x" in value and "y" in value:
        return int(value["x"]), int(value["y"])
    # String variants
    if isinstance(value, str):
        s = value.strip()
        if s.startswith("[") and s.endswith("]"):
            try:
                import json
                v = json.loads(s)
                if isinstance(v, list) and len(v) == 2:
                    return int(v[0]), int(v[1])
            except Exception:
                pass
        import re
        nums = re.findall(r"-?\d+", s)
        if len(nums) >= 2:
            return int(nums[0]), int(nums[1])
    raise ValueError("Location must be two integers [x,y]")


def _backend_from_env() -> InputBackend:
    """Construct a driver-level backend based on env settings.

    Prefers DLL backend and falls back to AHK when explicitly selected.
    Does not fall back to PyAutoGUI to guarantee driver-level injection.
    """
    preferred = (os.getenv("WINDOWS_MCP_INPUT_BACKEND", "ibsim-dll") or "").lower()
    driver = os.getenv("WINDOWS_MCP_INPUT_DRIVER", "AnyDriver")

    if preferred in ("ibsim-dll", "ibsim"):
        be = IBSimulatorDLLBackend(driver=driver)
        if be.info().ready:
            return be
        # If DLL path is not ready and user explicitly set ibsim-dll, keep it strict
        if preferred == "ibsim-dll":
            raise RuntimeError("IBSimulator DLL backend not ready. Check IbInputSimulator DLL files.")
        # Otherwise try AHK as a best-effort driver-level path
        ahk = IBSimulatorAHKBackend(driver=driver)
        if ahk.info().ready:
            return ahk
        raise RuntimeError("No driver-level backend is ready (DLL/AHK).")

    if preferred == "ibsim-ahk":
        ahk = IBSimulatorAHKBackend(driver=driver)
        if ahk.info().ready:
            return ahk
        raise RuntimeError("IBSimulator AHK backend not ready. Install AutoHotkey v2 and include files.")

    raise RuntimeError(f"Unsupported WINDOWS_MCP_INPUT_BACKEND={preferred}. Use 'ibsim-dll' or 'ibsim-ahk'.")


# Initialize backend and rate limiter at process start
backend: InputBackend = _backend_from_env()

rate = RateLimiter(
    RateConfig(
        mouse_move_hz=float(os.getenv("WINDOWS_MCP_RATE_MOVE_HZ", "120")),
        mouse_max_delta=int(os.getenv("WINDOWS_MCP_RATE_MAX_DELTA", "60")),
        mouse_smooth=float(os.getenv("WINDOWS_MCP_RATE_SMOOTH", "0.0")),
        clicks_per_sec=float(os.getenv("WINDOWS_MCP_RATE_CPS", "8.0")),
        keys_per_sec=float(os.getenv("WINDOWS_MCP_RATE_KPS", "12.0")),
    )
)


@asynccontextmanager
async def lifespan(app: FastMCP):
    try:
        yield
    finally:
        # No persistent resources to tear down
        pass


instructions = (
    "Driver-Level Input MCP exposes keyboard/mouse tools backed by IbInputSimulator (DLL/AHK).\n"
    "It guarantees OS-level injection without PyAutoGUI fallbacks.\n\n"
    "Parameter Reference:\n"
    "- Coordinates: pass [x,y], {x:..,y:..}, or string 'x,y'.\n"
    "- Buttons: 'left' | 'right' | 'middle'.\n"
    "- Shortcuts: single string like 'ctrl+c', 'win+r'.\n\n"
    "Examples:\n"
    "- Click: {\"loc\":[345,211],\"button\":\"left\",\"clicks\":2}\n"
    "- Type: {\"text\":\"你好\"}\n"
    "- Move: {\"to_loc\":[1280,720]}\n"
    "- Drag: {\"from_loc\":[500,400],\"to_loc\":[960,540]}\n"
    "- Shortcut: {\"shortcut\":\"ctrl+c\"}\n"
)

mcp = FastMCP(name="windows-input-mcp", instructions=instructions, lifespan=lifespan)


def _get_cursor_pos() -> tuple[int, int]:
    user32 = ctypes.windll.user32
    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
    pt = POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    return int(pt.x), int(pt.y)


def _set_clipboard_text(text: str) -> bool:
    CF_UNICODETEXT = 13
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    if not user32.OpenClipboard(None):
        return False
    try:
        if not user32.EmptyClipboard():
            return False
        data = text.encode('utf-16-le') + b"\x00\x00"
        hGlobal = kernel32.GlobalAlloc(0x0002, len(data))  # GMEM_MOVEABLE
        if not hGlobal:
            return False
        pchData = kernel32.GlobalLock(hGlobal)
        if not pchData:
            kernel32.GlobalFree(hGlobal)
            return False
        try:
            ctypes.memmove(pchData, data, len(data))
        finally:
            kernel32.GlobalUnlock(hGlobal)
        if not user32.SetClipboardData(CF_UNICODETEXT, hGlobal):
            kernel32.GlobalFree(hGlobal)
            return False
        return True
    finally:
        user32.CloseClipboard()


@mcp.tool(
    name="Input-Info",
    description="Return current backend info and rate limiter config."
)
def input_info() -> str:
    info = backend.info()
    rcfg = rate.cfg
    return (
        f"Backend: {info.name} ready={info.ready} details={info.details}\n"
        f"Rate: move_hz={rcfg.mouse_move_hz} max_delta={rcfg.mouse_max_delta} smooth={rcfg.mouse_smooth}; "
        f"cps={rcfg.clicks_per_sec} kps={rcfg.keys_per_sec}"
    )


@mcp.tool(
    name="Move-Tool",
    description="Move cursor to coordinates. Format: to_loc=[x,y] or 'x,y'. Uses stepwise movement under rate limits."
)
def move_tool(to_loc: list[int] | dict | str) -> str:
    tx, ty = _coerce_xy(to_loc)
    # Step toward target using RateLimiter
    for _ in range(3000):  # hard cap
        cx, cy = _get_cursor_pos()
        if (cx, cy) == (tx, ty):
            break
        nx, ny = rate.filter_target((cx, cy), (tx, ty))
        rate.sleep_until_ready("move")
        backend.move(nx, ny)
        if (nx, ny) == (tx, ty):
            # final snap if needed
            backend.move(tx, ty)
            break
    return f"Moved to ({tx},{ty})."


@mcp.tool(
    name="Click-Tool",
    description="Click at coordinates. Format: loc=[x,y], button='left|right|middle', clicks=int."
)
def click_tool(loc: list[int] | dict | str, button: Literal['left', 'right', 'middle'] = 'left', clicks: int = 1) -> str:
    x, y = _coerce_xy(loc)
    rate.sleep_until_ready("click")
    backend.click(x, y, button=button, clicks=int(max(1, clicks)))
    return f"{button} click x{int(max(1, clicks))} at ({x},{y})."


@mcp.tool(
    name="Drag-Tool",
    description="Drag from from_loc to to_loc. Format: from_loc=[x,y], to_loc=[x,y]."
)
def drag_tool(from_loc: list[int] | dict | str, to_loc: list[int] | dict | str) -> str:
    x1, y1 = _coerce_xy(from_loc)
    x2, y2 = _coerce_xy(to_loc)
    backend.drag(x1, y1, x2, y2)
    return f"Dragged from ({x1},{y1}) to ({x2},{y2})."


@mcp.tool(
    name="Type-Tool",
    description="Type text. method='unicode'|'clipboard'|'vk'. unicode uses KEYEVENTF_UNICODE; clipboard sets text then 'ctrl+v'; vk simulates per-char key events."
)
def type_tool(text: str, method: Literal['unicode', 'clipboard', 'vk'] = 'unicode', press_enter: bool = False) -> str:
    if not isinstance(text, str):
        raise ValueError("text must be a string")
    m = (method or 'unicode').lower()
    if m == 'clipboard':
        ok = _set_clipboard_text(text)
        if not ok:
            raise RuntimeError('Failed to set clipboard text')
        rate.sleep_until_ready('key')
        backend.hotkey('ctrl+v')
    elif m == 'vk':
        # Emit via key_down/up using keyboard mapping; safer for some games
        for ch in text:
            rate.sleep_until_ready('key')
            backend.key_down(ch)
            backend.key_up(ch)
    else:
        # default unicode path; chunk long strings
        for chunk in (text[i:i+64] for i in range(0, len(text), 64)):
            rate.sleep_until_ready('key')
            backend.send_text(chunk)
    if press_enter:
        rate.sleep_until_ready('key')
        backend.hotkey('enter')
    return f"Typed {len(text)} chars via {m}."


@mcp.tool(
    name="Shortcut-Tool",
    description="Send a keyboard shortcut, e.g., 'ctrl+c', 'win+r', 'shift+tab'."
)
def shortcut_tool(shortcut: str) -> str:
    if not isinstance(shortcut, str) or not shortcut:
        raise ValueError("shortcut must be a non-empty string, e.g., 'ctrl+c'")
    rate.sleep_until_ready("key")
    backend.hotkey(shortcut)
    return f"Shortcut sent: {shortcut}"


@mcp.tool(
    name="Key-Tool",
    description="Low-level key control. mode='tap'|'down'|'up'|'hold', key='w', times=1, interval_ms=40, hold_ms=0. Supports WASD and modifier keys."
)
def key_tool(
    mode: Literal['tap', 'down', 'up', 'hold'],
    key: str,
    times: int = 1,
    interval_ms: int = 40,
    hold_ms: int = 0,
) -> str:
    m = (mode or 'tap').lower()
    n = max(1, int(times))
    iv = max(0, int(interval_ms)) / 1000.0
    if m == 'down':
        rate.sleep_until_ready('key')
        backend.key_down(key)
        return f"Key down: {key}"
    if m == 'up':
        rate.sleep_until_ready('key')
        backend.key_up(key)
        return f"Key up: {key}"
    if m == 'hold':
        rate.sleep_until_ready('key')
        backend.key_down(key)
        try:
            import time
            time.sleep(max(0, int(hold_ms)) / 1000.0)
        finally:
            backend.key_up(key)
        return f"Key hold: {key} {max(0, int(hold_ms))}ms"
    # tap
    import time
    for _ in range(n):
        rate.sleep_until_ready('key')
        backend.key_down(key)
        time.sleep(max(0.0, min(0.25, iv)))
        backend.key_up(key)
        if iv > 0:
            time.sleep(iv)
    return f"Key tap: {key} x{n}"


@mcp.tool(
    name="Combo-Tool",
    description="Hold multiple keys together like human combos. keys=['shift','w'], hold_ms=600. Releases in reverse order."
)
def combo_tool(keys: list[str], hold_ms: int = 300) -> str:
    if not isinstance(keys, list) or not keys:
        raise ValueError("keys must be a non-empty list of strings")
    import time
    # Press in order
    for k in keys:
        rate.sleep_until_ready('key')
        backend.key_down(k)
    time.sleep(max(0, int(hold_ms)) / 1000.0)
    # Release in reverse
    for k in reversed(keys):
        backend.key_up(k)
    return f"Combo {keys} held {int(hold_ms)}ms"


@mcp.tool(
    name="Scroll-Tool",
    description="Scroll at optional coordinates. Format: loc?=[x,y], type='vertical|horizontal', direction, wheel_times=int."
)
def scroll_tool(
    loc: list[int] | dict | str | None = None,
    type: Literal['horizontal', 'vertical'] = 'vertical',
    direction: Literal['up', 'down', 'left', 'right'] = 'down',
    wheel_times: int = 1,
) -> str:
    if loc is not None:
        x, y = _coerce_xy(loc)
        backend.move(x, y)
    # Use backend-provided scroll to keep driver/OS-level semantics
    backend.scroll(int(max(1, wheel_times)), type, direction)
    pos = ""
    if loc is not None:
        pos = f" at ({x},{y})"
    return f"Scrolled {type} {direction} x{int(max(1, wheel_times))}{pos}."


@mcp.tool(
    name="Input-RateLimiter-Config",
    description="Configure rate limiting: move_hz, max_delta, smooth, cps, kps."
)
def rate_config(
    move_hz: float | None = None,
    max_delta: int | None = None,
    smooth: float | None = None,
    cps: float | None = None,
    kps: float | None = None,
) -> str:
    def _num(v, cur, cast):
        if v is None:
            return cur
        try:
            return cast(v)
        except Exception:
            return cur
    cfg = rate.cfg
    cfg.mouse_move_hz = max(15.0, min(480.0, _num(move_hz, cfg.mouse_move_hz, float)))
    cfg.mouse_max_delta = max(1, _num(max_delta, cfg.mouse_max_delta, int))
    cfg.mouse_smooth = max(0.0, min(0.98, _num(smooth, cfg.mouse_smooth, float)))
    cfg.clicks_per_sec = max(1.0, min(60.0, _num(cps, cfg.clicks_per_sec, float)))
    cfg.keys_per_sec = max(1.0, min(60.0, _num(kps, cfg.keys_per_sec, float)))
    return (
        f"rate(move_hz={cfg.mouse_move_hz}, max_delta={cfg.mouse_max_delta}, smooth={cfg.mouse_smooth}, "
        f"cps={cfg.clicks_per_sec}, kps={cfg.keys_per_sec})"
    )


@click.command()
@click.option(
    "--transport",
    help="The transport layer used by the MCP server.",
    type=click.Choice(["stdio", "sse", "streamable-http"]),
    default="stdio",
)
@click.option(
    "--host",
    help="Host to bind the SSE/Streamable HTTP server.",
    default="localhost",
    type=str,
    show_default=True,
)
@click.option(
    "--port",
    help="Port to bind the SSE/Streamable HTTP server.",
    default=8001,
    type=int,
    show_default=True,
)
def main(transport, host, port):
    if transport == "stdio":
        mcp.run()
    else:
        mcp.run(transport=transport, host=host, port=port)


if __name__ == "__main__":
    main()
