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
    "- Desktop-Info: {}\n"
    "- Window-Info: {} (returns active window rect/title/class)\n"
    "- Windows-List: {\"query\":\"chrome\", \"only_visible\":true, \"include_minimized\":true, \"limit\":20}\n"
    "- Windows-List (children): {\"parent_hwnd\": 131072, \"only_visible\":false}\n"
)
instructions = instructions + "- Windows-Select: {\"query\":\"notepad\"} or {\"index\":0}\n"

mcp = FastMCP(name="windows-input-mcp", instructions=instructions, lifespan=lifespan)


# --- MCP Prompts & Resources -------------------------------------------------
# Some MCP clients (e.g., Claude Desktop) surface prompts/resources explicitly.
# We register a few practical prompts and read-only resources to improve UX.

def _register_prompts_and_resources() -> None:
    """Register MCP prompts and resources with graceful fallback.

    Works across FastMCP versions by checking for decorators/methods at runtime.
    No-ops if the feature is unavailable in the installed FastMCP.
    """
    # Prompts: provide guided starters for common actions
    try:
        if hasattr(mcp, "prompt"):
            # Click at point
            @mcp.prompt(name="Click-At", description="Click at coordinates using Click-Tool.")
            def prompt_click(loc: str, button: str = "left", clicks: int = 1):
                return [
                    {"role": "system", "content": "Use Click-Tool for driver-level clicking."},
                    {"role": "user", "content": f"Call Click-Tool with loc='{loc}', button='{button}', clicks={int(clicks)}"},
                ]

            # Type text
            @mcp.prompt(name="Type-Text", description="Type text via driver-level injection.")
            def prompt_type(text: str, method: str = "unicode"):
                return [
                    {"role": "system", "content": "Use Type-Tool for text input."},
                    {"role": "user", "content": f"Call Type-Tool with text='{text}', method='{method}'"},
                ]

            # Shortcut
            @mcp.prompt(name="Send-Shortcut", description="Send a keyboard shortcut like 'ctrl+c' or 'win+r'.")
            def prompt_shortcut(shortcut: str):
                return [
                    {"role": "system", "content": "Use Shortcut-Tool for combos."},
                    {"role": "user", "content": f"Call Shortcut-Tool with shortcut='{shortcut}'"},
                ]

            # Activate window by query or index
            @mcp.prompt(name="Activate-Window", description="List/select a window and activate it.")
            def prompt_activate_window(query: str = "", index: int | None = None):
                idx_part = " first match" if index is None else f" index={int(index)}"
                return [
                    {"role": "system", "content": "Use Windows-List then Windows-Select, then Windows-Activate."},
                    {"role": "user", "content": f"List windows with query='{query}'. Select{idx_part}, then activate."},
                ]

            # Drag operation
            @mcp.prompt(name="Drag-From-To", description="Drag from one point to another.")
            def prompt_drag(from_loc: str, to_loc: str):
                return [
                    {"role": "system", "content": "Use Drag-Tool for driver-level dragging."},
                    {"role": "user", "content": f"Call Drag-Tool with from_loc='{from_loc}', to_loc='{to_loc}'"},
                ]
        elif hasattr(mcp, "add_prompt"):
            # Fallback builder-style registration (single-message prompts)
            try:
                mcp.add_prompt(
                    name="Click-At",
                    description="Click at coordinates using Click-Tool.",
                    arguments=[
                        {"name": "loc", "type": "string", "required": True},
                        {"name": "button", "type": "string", "required": False},
                        {"name": "clicks", "type": "number", "required": False},
                    ],
                    messages=[
                        {"role": "system", "content": "Use Click-Tool."},
                        {"role": "user", "content": "Call Click-Tool with loc={{loc}} button={{button}} clicks={{clicks}}"},
                    ],
                )
            except Exception:
                pass
            try:
                mcp.add_prompt(
                    name="Type-Text",
                    description="Type text via driver-level injection.",
                    arguments=[
                        {"name": "text", "type": "string", "required": True},
                        {"name": "method", "type": "string", "required": False},
                    ],
                    messages=[
                        {"role": "system", "content": "Use Type-Tool."},
                        {"role": "user", "content": "Call Type-Tool with text={{text}} method={{method}}"},
                    ],
                )
            except Exception:
                pass
    except Exception as e:
        logger.debug("Prompt registration skipped: %s", e)

    # Resources: read-only helpful views powered by existing tools
    try:
        if hasattr(mcp, "resource"):
            @mcp.resource(uri="mcp://windows/desktop-info", name="Desktop Info", description="Virtual screen and monitors")
            def res_desktop_info() -> str:
                return desktop_info()

            @mcp.resource(uri="mcp://windows/active-window", name="Active Window", description="Title/class/rect of foreground window")
            def res_active_window() -> str:
                return window_info()

            @mcp.resource(uri="mcp://windows/rate", name="Input Rate", description="Backend + rate limiter settings")
            def res_rate() -> str:
                return input_info()

            @mcp.resource(uri="mcp://windows/instructions", name="Server Instructions", description="Usage tips and examples")
            def res_instructions() -> str:
                return instructions

            @mcp.resource(uri="mcp://windows/env", name="Runtime Env", description="Relevant WINDOWS_MCP_* environment variables")
            def res_env() -> str:
                keys = [
                    "WINDOWS_MCP_INPUT_BACKEND",
                    "WINDOWS_MCP_INPUT_DRIVER",
                    "WINDOWS_MCP_RATE_MOVE_HZ",
                    "WINDOWS_MCP_RATE_MAX_DELTA",
                    "WINDOWS_MCP_RATE_SMOOTH",
                    "WINDOWS_MCP_RATE_CPS",
                    "WINDOWS_MCP_RATE_KPS",
                    "WINDOWS_INPUT_LOG_LEVEL",
                ]
                lines = [f"{k}={os.getenv(k, '')}" for k in keys]
                return "\n".join(lines)
        elif hasattr(mcp, "add_resource"):
            try:
                mcp.add_resource(uri="mcp://windows/desktop-info", name="Desktop Info", mimeType="text/plain", value=desktop_info())
                mcp.add_resource(uri="mcp://windows/active-window", name="Active Window", mimeType="text/plain", value=window_info())
                mcp.add_resource(uri="mcp://windows/rate", name="Input Rate", mimeType="text/plain", value=input_info())
                mcp.add_resource(uri="mcp://windows/instructions", name="Server Instructions", mimeType="text/plain", value=instructions)
            except Exception:
                pass
    except Exception as e:
        logger.debug("Resource registration skipped: %s", e)


# Register at import time so prompts/resources are visible before run()
_register_prompts_and_resources()


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


def _parse_hwnd(value) -> int:
    """Parse hwnd from int or hex/decimal string (e.g., 123456, '0x001234AB')."""
    if isinstance(value, int):
        return int(value)
    if isinstance(value, str):
        s = value.strip().lower()
        if s.startswith('0x'):
            return int(s, 16)
        return int(s)
    raise ValueError("Invalid hwnd format; use int or hex string like '0x001234AB'")


def _coerce_wh(value) -> tuple[int, int]:
    if isinstance(value, (list, tuple)) and len(value) == 2:
        return int(value[0]), int(value[1])
    if isinstance(value, dict) and 'w' in value and 'h' in value:
        return int(value['w']), int(value['h'])
    if isinstance(value, str):
        import re
        nums = re.findall(r'-?\\d+', value)
        if len(nums) >= 2:
            return int(nums[0]), int(nums[1])
    raise ValueError("size must be [w,h] or {w,h}")


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
    name="Desktop-Info",
    description="Get virtual screen origin/size and monitor count."
)
def desktop_info() -> str:
    user32 = ctypes.windll.user32
    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79
    SM_CMONITORS = 80
    left = int(user32.GetSystemMetrics(SM_XVIRTUALSCREEN))
    top = int(user32.GetSystemMetrics(SM_YVIRTUALSCREEN))
    width = int(user32.GetSystemMetrics(SM_CXVIRTUALSCREEN))
    height = int(user32.GetSystemMetrics(SM_CYVIRTUALSCREEN))
    monitors = int(user32.GetSystemMetrics(SM_CMONITORS))
    return (
        f"VirtualScreen: left={left}, top={top}, width={width}, height={height}, monitors={monitors}"
    )


@mcp.tool(
    name="Window-Info",
    description="Get active window title, class, and rect [l,t,r,b]."
)
def window_info() -> str:
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return "No active window."
    # Title
    GetWindowTextLengthW = user32.GetWindowTextLengthW
    GetWindowTextW = user32.GetWindowTextW
    length = int(GetWindowTextLengthW(hwnd))
    buf = ctypes.create_unicode_buffer(length + 1) if length > 0 else ctypes.create_unicode_buffer(1)
    GetWindowTextW(hwnd, buf, len(buf))
    title = buf.value
    # Class
    bufc = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, bufc, 256)
    cls = bufc.value
    # Rect
    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
    rc = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rc))
    l, t, r, b = int(rc.left), int(rc.top), int(rc.right), int(rc.bottom)
    w, h = r - l, b - t
    return (
        f"ActiveWindow: hwnd={int(hwnd)} title='{title}' class='{cls}' rect=[{l},{t},{r},{b}] size=[{w},{h}]"
    )


@mcp.tool(
    name="Windows-Activate",
    description="Activate/bring window to front. Args: hwnd=int|hex, show='restore|minimize|maximize|show' (optional), topmost=bool (optional)."
)
def windows_activate(hwnd: int | str, show: str | None = None, topmost: bool | None = None) -> str:
    user32 = ctypes.windll.user32
    target = _parse_hwnd(hwnd)

    SW_RESTORE = 9
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_MAXIMIZE = 3

    # Show state first (if requested)
    if show:
        m = show.strip().lower()
        cmd = None
        if m == 'restore': cmd = SW_RESTORE
        elif m == 'show': cmd = SW_SHOW
        elif m == 'minimize': cmd = SW_MINIMIZE
        elif m == 'maximize': cmd = SW_MAXIMIZE
        if cmd is not None:
            user32.ShowWindow(wintypes.HWND(target), int(cmd))

    # Try SetForegroundWindow; fallback to AttachThreadInput method
    ok = bool(user32.SetForegroundWindow(wintypes.HWND(target)))
    if not ok:
        fg = user32.GetForegroundWindow()
        fg_tid = ctypes.c_ulong(0)
        tgt_tid = ctypes.c_ulong(0)
        user32.GetWindowThreadProcessId(fg, ctypes.byref(fg_tid))
        user32.GetWindowThreadProcessId(wintypes.HWND(target), ctypes.byref(tgt_tid))
        try:
            user32.AttachThreadInput(fg_tid.value, tgt_tid.value, True)
            user32.BringWindowToTop(wintypes.HWND(target))
            ok = bool(user32.SetForegroundWindow(wintypes.HWND(target)))
        finally:
            user32.AttachThreadInput(fg_tid.value, tgt_tid.value, False)

    # Optional topmost toggle
    if topmost is not None:
        HWND_TOPMOST = ctypes.c_void_p(-1)
        HWND_NOTOPMOST = ctypes.c_void_p(-2)
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_SHOWWINDOW = 0x0040
        user32.SetWindowPos(
            wintypes.HWND(target),
            HWND_TOPMOST if bool(topmost) else HWND_NOTOPMOST,
            0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
        )

    # Return final rect
    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
    rc = RECT(); user32.GetWindowRect(wintypes.HWND(target), ctypes.byref(rc))
    l, t, r, b = int(rc.left), int(rc.top), int(rc.right), int(rc.bottom)
    w, h = r - l, b - t
    return f"Activate {'OK' if ok else 'TRY'} hwnd=0x{target:08X} rect=[{l},{t},{r},{b}] size=[{w},{h}] topmost={'on' if topmost else 'unchanged' if topmost is None else 'off'}"


@mcp.tool(
    name="Windows-SetPos",
    description="Move/resize window via SetWindowPos. Args: hwnd, loc=[x,y]?, size=[w,h]?, z='topmost|notopmost|top|bottom|nochange' (optional)."
)
def windows_setpos(
    hwnd: int | str,
    loc: list[int] | dict | str | None = None,
    size: list[int] | dict | str | None = None,
    z: str | None = None,
) -> str:
    user32 = ctypes.windll.user32
    target = _parse_hwnd(hwnd)
    x = y = 0
    w = h = 0
    flags = 0
    SWP_NOSIZE = 0x0001
    SWP_NOMOVE = 0x0002
    SWP_NOZORDER = 0x0004
    SWP_SHOWWINDOW = 0x0040

    if loc is None:
        flags |= SWP_NOMOVE
    else:
        x, y = _coerce_xy(loc)
    if size is None:
        flags |= SWP_NOSIZE
    else:
        w, h = _coerce_wh(size)

    insert_after = ctypes.c_void_p(0)
    if z:
        zmap = {
            'topmost': ctypes.c_void_p(-1),
            'notopmost': ctypes.c_void_p(-2),
            'top': ctypes.c_void_p(0),
            'bottom': ctypes.c_void_p(1),
        }
        insert_after = zmap.get(z.strip().lower(), ctypes.c_void_p(0))
    else:
        flags |= SWP_NOZORDER

    flags |= SWP_SHOWWINDOW
    ok = bool(user32.SetWindowPos(wintypes.HWND(target), insert_after, int(x), int(y), int(w), int(h), int(flags)))

    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
    rc = RECT(); user32.GetWindowRect(wintypes.HWND(target), ctypes.byref(rc))
    l, t, r, b = int(rc.left), int(rc.top), int(rc.right), int(rc.bottom)
    return f"SetPos {'OK' if ok else 'FAIL'} hwnd=0x{target:08X} rect=[{l},{t},{r},{b}]"


@mcp.tool(
    name="Windows-Close",
    description="Send WM_CLOSE to a window. Args: hwnd=int|hex."
)
def windows_close(hwnd: int | str) -> str:
    user32 = ctypes.windll.user32
    target = _parse_hwnd(hwnd)
    WM_CLOSE = 0x0010
    ok = bool(user32.PostMessageW(wintypes.HWND(target), WM_CLOSE, 0, 0))
    return f"Close {'OK' if ok else 'FAIL'} hwnd=0x{target:08X}"


@mcp.tool(
    name="Windows-List",
    description="Enumerate windows. Top-level via EnumWindows or children via EnumChildWindows when parent_hwnd is set. Filters: query, only_visible, include_minimized, include_cloaked, limit."
)
def windows_list(
    query: str | None = None,
    only_visible: bool = True,
    include_minimized: bool = True,
    include_cloaked: bool = True,
    limit: int | None = None,
    parent_hwnd: int | None = None,
) -> str:
    user32 = ctypes.windll.user32

    EnumWindows = user32.EnumWindows
    IsWindowVisible = user32.IsWindowVisible
    IsIconic = user32.IsIconic
    GetWindowTextLengthW = user32.GetWindowTextLengthW
    GetWindowTextW = user32.GetWindowTextW
    GetClassNameW = user32.GetClassNameW
    GetWindowRect = user32.GetWindowRect
    GetWindowThreadProcessId = user32.GetWindowThreadProcessId
    EnumChildWindows = user32.EnumChildWindows

    # DWM cloaking (occluded/hidden by OS)
    cloaked_attr = 14  # DWMWA_CLOAKED
    try:
        dwmapi = ctypes.windll.dwmapi
        DwmGetWindowAttribute = dwmapi.DwmGetWindowAttribute
        def _is_cloaked(hwnd):
            val = ctypes.c_int(0)
            try:
                DwmGetWindowAttribute(hwnd, cloaked_attr, ctypes.byref(val), ctypes.sizeof(val))
                return bool(val.value)
            except Exception:
                return False
    except Exception:
        def _is_cloaked(hwnd):
            return False

    results: list[dict] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, ctypes.c_void_p)
    def _enum_proc(hwnd, lparam):
        try:
            visible = bool(IsWindowVisible(hwnd))
            minimized = bool(IsIconic(hwnd))
            cloaked = _is_cloaked(hwnd)
            if only_visible and not visible:
                return True
            if minimized and not include_minimized:
                return True
            if cloaked and not include_cloaked:
                return True
            # Title
            length = int(GetWindowTextLengthW(hwnd))
            tbuf = ctypes.create_unicode_buffer(length + 1) if length > 0 else ctypes.create_unicode_buffer(1)
            GetWindowTextW(hwnd, tbuf, len(tbuf))
            title = tbuf.value
            # Class
            cbuf = ctypes.create_unicode_buffer(256)
            GetClassNameW(hwnd, cbuf, 256)
            cls = cbuf.value
            # Rect
            class RECT(ctypes.Structure):
                _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
            rc = RECT()
            GetWindowRect(hwnd, ctypes.byref(rc))
            # PID
            pid = ctypes.c_ulong()
            GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

            item = {
                'hwnd': int(hwnd),
                'pid': int(pid.value),
                'class': cls,
                'title': title,
                'left': int(rc.left), 'top': int(rc.top), 'right': int(rc.right), 'bottom': int(rc.bottom),
                'visible': visible,
                'minimized': minimized,
                'cloaked': cloaked,
            }

            q = (query or '').strip().lower()
            if q:
                if q not in (title or '').lower() and q not in (cls or '').lower() and q not in str(item['pid']):
                    return True

            results.append(item)
            if limit is not None and len(results) >= int(limit):
                return False  # stop enumeration
        except Exception:
            pass
        return True

    if parent_hwnd is not None and int(parent_hwnd) != 0:
        EnumChildWindows(wintypes.HWND(int(parent_hwnd)), _enum_proc, 0)
    else:
        EnumWindows(_enum_proc, 0)

    lines = [
        "idx  hwnd        pid   V M C  class                title                          rect[l,t,r,b]  size[w,h]",
        "---- ----------- ----- - - - -------------------- ------------------------------ -------------- ----------",
    ]
    for i, it in enumerate(results):
        l, t, r, b = it['left'], it['top'], it['right'], it['bottom']
        w, h = r - l, b - t
        lines.append(
            f"{i:>3}  0x{it['hwnd']:08X} {it['pid']:>5} "
            f"{ 'Y' if it['visible'] else '-' } { 'Y' if it['minimized'] else '-' } { 'Y' if it['cloaked'] else '-' } "
            f"{it['class'][:20]:<20} {(it['title'] or '')[:30]:<30} [{l},{t},{r},{b}] {w}x{h}"
        )
    if not results:
        lines.append("(no windows matched)")
    return "\n".join(lines)


@mcp.tool(
    name="Windows-Select",
    description="Select a window and return hwnd using the same filters as Windows-List. Prefer index when provided; otherwise pick the first match."
)
def windows_select(
    query: str | None = None,
    only_visible: bool = True,
    include_minimized: bool = True,
    include_cloaked: bool = True,
    index: int | None = None,
    parent_hwnd: int | None = None,
) -> str:
    user32 = ctypes.windll.user32

    EnumWindows = user32.EnumWindows
    IsWindowVisible = user32.IsWindowVisible
    IsIconic = user32.IsIconic
    GetWindowTextLengthW = user32.GetWindowTextLengthW
    GetWindowTextW = user32.GetWindowTextW
    GetClassNameW = user32.GetClassNameW
    GetWindowRect = user32.GetWindowRect
    GetWindowThreadProcessId = user32.GetWindowThreadProcessId
    EnumChildWindows = user32.EnumChildWindows

    # DWM cloaking
    cloaked_attr = 14  # DWMWA_CLOAKED
    try:
        dwmapi = ctypes.windll.dwmapi
        DwmGetWindowAttribute = dwmapi.DwmGetWindowAttribute
        def _is_cloaked(hwnd):
            val = ctypes.c_int(0)
            try:
                DwmGetWindowAttribute(hwnd, cloaked_attr, ctypes.byref(val), ctypes.sizeof(val))
                return bool(val.value)
            except Exception:
                return False
    except Exception:
        def _is_cloaked(hwnd):
            return False

    results: list[dict] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, ctypes.c_void_p)
    def _enum_proc(hwnd, lparam):
        try:
            visible = bool(IsWindowVisible(hwnd))
            minimized = bool(IsIconic(hwnd))
            cloaked = _is_cloaked(hwnd)
            if only_visible and not visible:
                return True
            if minimized and not include_minimized:
                return True
            if cloaked and not include_cloaked:
                return True
            length = int(GetWindowTextLengthW(hwnd))
            tbuf = ctypes.create_unicode_buffer(length + 1) if length > 0 else ctypes.create_unicode_buffer(1)
            GetWindowTextW(hwnd, tbuf, len(tbuf))
            title = tbuf.value
            cbuf = ctypes.create_unicode_buffer(256)
            GetClassNameW(hwnd, cbuf, 256)
            cls = cbuf.value
            class RECT(ctypes.Structure):
                _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
            rc = RECT(); GetWindowRect(hwnd, ctypes.byref(rc))
            pid = ctypes.c_ulong(); GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            item = {
                'hwnd': int(hwnd), 'pid': int(pid.value), 'class': cls, 'title': title,
                'left': int(rc.left), 'top': int(rc.top), 'right': int(rc.right), 'bottom': int(rc.bottom),
                'visible': visible, 'minimized': minimized, 'cloaked': cloaked,
            }
            q = (query or '').strip().lower()
            if q:
                if q not in (title or '').lower() and q not in (cls or '').lower() and q not in str(item['pid']):
                    return True
            results.append(item)
        except Exception:
            pass
        return True

    if parent_hwnd is not None and int(parent_hwnd) != 0:
        EnumChildWindows(wintypes.HWND(int(parent_hwnd)), _enum_proc, 0)
    else:
        EnumWindows(_enum_proc, 0)

    if not results:
        return "No windows matched."

    if index is not None:
        i = int(index)
        if i < 0 or i >= len(results):
            return f"Index out of range (0..{len(results)-1})."
        it = results[i]
    else:
        it = results[0]

    l, t, r, b = it['left'], it['top'], it['right'], it['bottom']
    w, h = r - l, b - t
    return (
        f"Selected hwnd=0x{it['hwnd']:08X} pid={it['pid']} visible={it['visible']} minimized={it['minimized']} cloaked={it['cloaked']}"
        f" class='{it['class']}' title='{it['title']}' rect=[{l},{t},{r},{b}] size=[{w},{h}]"
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
        # Default unicode path - send entire text to backend
        # Backend now handles character-by-character sending with proper delays
        # No need to chunk here, backend.send_text has built-in delays
        rate.sleep_until_ready('key')
        backend.send_text(text)
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
