from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from shutil import which
import logging
import subprocess
import tempfile
import textwrap
import os
import ctypes
from ctypes import wintypes


logger = logging.getLogger(__name__)


@dataclass
class BackendInfo:
    name: str
    ready: bool
    details: str = ""


def _find_ahk_exe() -> str | None:
    env = os.getenv("AUTOHOTKEY_EXE")
    if env and Path(env).exists():
        return env
    for name in ("AutoHotkey64.exe", "AutoHotkeyU64.exe", "AutoHotkey.exe", "autohotkey.exe"):
        exe = which(name)
        if exe:
            return exe
    pf = os.environ.get("ProgramFiles", r"C:\\Program Files")
    pfx86 = os.environ.get("ProgramFiles(x86)", r"C:\\Program Files (x86)")
    candidates = [
        Path(pf) / "AutoHotkey" / "v2" / "AutoHotkey64.exe",
        Path(pf) / "AutoHotkey" / "AutoHotkey64.exe",
        Path(pf) / "AutoHotkey" / "v2" / "AutoHotkeyU64.exe",
        Path(pf) / "AutoHotkey" / "AutoHotkeyU64.exe",
        Path(pf) / "AutoHotkey" / "v2" / "AutoHotkey.exe",
        Path(pf) / "AutoHotkey" / "AutoHotkey.exe",
        Path(pfx86) / "AutoHotkey" / "v2" / "AutoHotkey64.exe",
        Path(pfx86) / "AutoHotkey" / "AutoHotkey64.exe",
        Path(pfx86) / "AutoHotkey" / "v2" / "AutoHotkeyU64.exe",
        Path(pfx86) / "AutoHotkey" / "AutoHotkeyU64.exe",
        Path(pfx86) / "AutoHotkey" / "v2" / "AutoHotkey.exe",
        Path(pfx86) / "AutoHotkey" / "AutoHotkey.exe",
    ]
    for c in candidates:
        if c.exists():
            return str(c)
    return None


def _ib_candidate_dirs() -> list[Path]:
    here = Path(__file__).resolve()
    env = os.getenv("IBSIM_DIR")
    out: list[Path] = []
    if env:
        out.append(Path(env))
    # local vendor in this project (same folder as this file)
    out.append(here.parent / "IbInputSimulator" / "Binding.AHK2")
    out.append(here.parent / "IbInputSimulator")
    # project root candidate (monorepo layout)
    out.append(here.parents[1] / "IbInputSimulator" / "Binding.AHK2")
    out.append(here.parents[1] / "IbInputSimulator")
    # parent repo layout fallback
    out.append(here.parents[2] / "IbInputSimulator" / "Binding.AHK2")
    out.append(here.parents[2] / "IbInputSimulator")
    # cwd fallback
    out.append(Path.cwd() / "IbInputSimulator" / "Binding.AHK2")
    out.append(Path.cwd() / "IbInputSimulator")
    return out


def _ib_ahk_include_path() -> Path | None:
    for d in _ib_candidate_dirs():
        p = d / "IbInputSimulator.ahk" if d.name == "Binding.AHK2" else d / "Binding.AHK2" / "IbInputSimulator.ahk"
        if p.exists():
            return p
    return None


def _ib_dll_path() -> Path | None:
    for d in _ib_candidate_dirs():
        p = d / "IbInputSimulator.dll" if d.name == "Binding.AHK2" else d / "Binding.AHK2" / "IbInputSimulator.dll"
        if p.exists():
            return p
    return None


class InputBackend:
    def info(self) -> BackendInfo:  # pragma: no cover
        raise NotImplementedError

    # Mouse
    def move(self, x: int, y: int):  # pragma: no cover
        raise NotImplementedError

    def click(self, x: int, y: int, button: str = "left", clicks: int = 1):  # pragma: no cover
        raise NotImplementedError

    def drag(self, x1: int, y1: int, x2: int, y2: int):  # pragma: no cover
        raise NotImplementedError

    def scroll(self, wheel_times: int, type: str = "vertical", direction: str = "down"):  # pragma: no cover
        raise NotImplementedError

    # Keyboard
    def send_text(self, text: str):  # pragma: no cover
        raise NotImplementedError

    def hotkey(self, combo: str):  # e.g., "ctrl+c"
        raise NotImplementedError

    def key_down(self, key: str):  # pragma: no cover
        raise NotImplementedError

    def key_up(self, key: str):  # pragma: no cover
        raise NotImplementedError

    def tap(self, key: str):  # pragma: no cover
        self.key_down(key)
        self.key_up(key)


class IBSimulatorAHKBackend(InputBackend):
    def __init__(self, driver: str = "AnyDriver"):
        self._ahk = _find_ahk_exe()
        self._inc = _ib_ahk_include_path()
        self._dll = _ib_dll_path()
        self._driver = driver
        self._ready = bool(self._ahk and self._inc and self._inc.exists() and self._dll and self._dll.exists())

    def info(self) -> BackendInfo:
        details = f"ahk={self._ahk}, include={self._inc}, dll={self._dll}"
        return BackendInfo("IBSimulatorAHK", self._ready, details)

    def _run(self, body: str) -> int:
        if not self._ready:
            return 1
        inc_path = str(self._inc)
        dll_path = str(self._dll)
        hdr = textwrap.dedent(f"""
        #Requires AutoHotkey v2.0
        #NoTrayIcon
        SetBatchLines -1
        #DllLoad "*i {dll_path}"
        #Include "{inc_path}"
        try {{
            IbSendInit("{self._driver}")
        }} catch e {{
            try {{ IbSendInit("SendInput") }} catch e2 {{ SendMode "Input" }}
        }}
        CoordMode "Mouse", "Screen"
        CoordMode "Pixel", "Screen"
        """)
        script = hdr + "\n" + body + "\nExitApp\n"
        with tempfile.NamedTemporaryFile(prefix="ibsim_", suffix=".ahk", delete=False) as tf:
            tf.write(script.encode("utf-8"))
            tf_path = tf.name
        try:
            proc = subprocess.run([self._ahk, "/ErrorStdOut", tf_path], timeout=6)
            return proc.returncode
        finally:
            try:
                os.remove(tf_path)
            except Exception:
                pass

    def move(self, x: int, y: int):
        body = (
            "try {\n"
            f"    IbMouseMove {x}, {y}, 0\n"
            "} catch e {\n"
            f"    MouseMove {x}, {y}, 0\n"
            "}"
        )
        self._run(body)

    def click(self, x: int, y: int, button: str = "left", clicks: int = 1):
        btn = button.lower()
        body = (
            "try {\n"
            f"    IbMouseClick \"{btn}\", {x}, {y}, {clicks}, 0\n"
            "} catch e {\n"
            f"    MouseClick \"{btn}\", {x}, {y}, {clicks}, 0\n"
            "}"
        )
        self._run(body)

    def drag(self, x1: int, y1: int, x2: int, y2: int):
        body = (
            "try {\n"
            f"    IbMouseClickDrag \"left\", {x1}, {y1}, {x2}, {y2}, 0\n"
            "} catch e {\n"
            f"    MouseClickDrag \"left\", {x1}, {y1}, {x2}, {y2}, 0\n"
            "}"
        )
        self._run(body)

    def scroll(self, wheel_times: int, type: str = "vertical", direction: str = "down"):
        n = max(1, int(wheel_times))
        t = (type or "vertical").lower()
        d = (direction or "down").lower()
        key = None
        if t == "vertical":
            if d == "up": key = "WheelUp"
            elif d == "down": key = "WheelDown"
        elif t == "horizontal":
            if d == "left": key = "WheelLeft"
            elif d == "right": key = "WheelRight"
        if not key:
            return
        body = (
            "try {\n"
            "    IbSendMode(1)\n"
            f"    Send(\"{{{key} {n}}}\")\n"
            "    IbSendMode(0)\n"
            "} catch e {\n"
            "    SendMode \"Input\"\n"
            f"    Send(\"{{{key} {n}}}\")\n"
            "}"
        )
        self._run(body)

    def send_text(self, text: str):
        # Send character by character with delay for reliability
        # SetKeyDelay adds delay between key down and up
        # Sleep adds delay between characters
        body_parts = [
            "try {",
            "    IbSendMode(1)",
            "    SetKeyDelay 3, 20",  # 3ms down-to-up, 20ms between keys
        ]

        for ch in text:
            esc_ch = ch.replace('"', '""')
            body_parts.append(f"    Send(\"{{Text}}\" . \"{esc_ch}\")")

        body_parts.extend([
            "    IbSendMode(0)",
            "} catch e {",
            "    SendMode \"Input\"",
            "    SetKeyDelay 3, 20",
        ])

        for ch in text:
            esc_ch = ch.replace('"', '""')
            body_parts.append(f"    Send(\"{{Text}}\" . \"{esc_ch}\")")

        body_parts.append("}")

        body = "\n".join(body_parts)
        self._run(body)

    def hotkey(self, combo: str):
        parts = [p.strip().lower() for p in combo.split("+") if p.strip()]
        mod_map = {"ctrl": "^", "alt": "!", "shift": "+", "win": "#"}
        key_name_map = {
            "enter": "Enter", "return": "Enter", "backspace": "Backspace",
            "delete": "Delete", "insert": "Insert", "tab": "Tab", "esc": "Escape",
            "escape": "Escape", "space": "Space", "home": "Home", "end": "End",
            "pgup": "PgUp", "pageup": "PgUp", "pgdn": "PgDn", "pagedown": "PgDn",
            "up": "Up", "down": "Down", "left": "Left", "right": "Right",
        }
        mods = []
        key = None
        for p in parts:
            if p in mod_map:
                mods.append(mod_map[p])
            elif len(p) == 1:
                key = p
            else:
                key = "{" + key_name_map.get(p, p.capitalize()) + "}"
        if key is None:
            back_map = {"^": "{Ctrl}", "!": "{Alt}", "+": "{Shift}", "#": "{LWin}"}
            key = back_map.get(mods[-1], "") if mods else ""
            mods = mods[:-1] if mods else []
        ahk_seq = "".join(mods) + (key or "")
        body = (
            "try {\n"
            f"    IbSend(\"{ahk_seq}\")\n"
            "} catch e {\n"
            f"    Send(\"{ahk_seq}\")\n"
            "}"
        )
        self._run(body)

    def _ahk_key_name(self, key: str) -> str:
        p = key.strip().lower()
        if p.startswith('vk') and len(p) >= 4:
            return 'vk' + p[2:].lstrip('_').upper()
        key_name_map = {
            'enter': 'Enter', 'return': 'Enter', 'backspace': 'Backspace', 'bs': 'Backspace',
            'tab': 'Tab', 'esc': 'Escape', 'escape': 'Escape', 'space': 'Space',
            'home': 'Home', 'end': 'End', 'pgup': 'PgUp', 'pageup': 'PgUp', 'pgdn': 'PgDn', 'pagedown': 'PgDn',
            'up': 'Up', 'down': 'Down', 'left': 'Left', 'right': 'Right',
            'shift': 'Shift', 'lshift': 'LShift', 'rshift': 'RShift',
            'ctrl': 'Ctrl', 'control': 'Ctrl', 'lctrl': 'LCtrl', 'rctrl': 'RCtrl',
            'alt': 'Alt', 'lalt': 'LAlt', 'ralt': 'RAlt',
            'win': 'LWin', 'lwin': 'LWin', 'rwin': 'RWin', 'apps': 'AppsKey', 'menu': 'AppsKey',
        }
        if p in key_name_map:
            return key_name_map[p]
        if p.startswith('f') and p[1:].isdigit():
            n = int(p[1:])
            if 1 <= n <= 24:
                return f'F{n}'
        return p if len(p) == 1 else p.capitalize()

    def key_down(self, key: str):
        k = self._ahk_key_name(key)
        body = (
            "try {\n"
            "    IbSendMode(1)\n"
            f"    Send(\"{{{k} down}}\")\n"
            "    IbSendMode(0)\n"
            "} catch e {\n"
            "    SendMode \"Input\"\n"
            f"    Send(\"{{{k} down}}\")\n"
            "}"
        )
        self._run(body)

    def key_up(self, key: str):
        k = self._ahk_key_name(key)
        body = (
            "try {\n"
            "    IbSendMode(1)\n"
            f"    Send(\"{{{k} up}}\")\n"
            "    IbSendMode(0)\n"
            "} catch e {\n"
            "    SendMode \"Input\"\n"
            f"    Send(\"{{{k} up}}\")\n"
            "}"
        )
        self._run(body)


class IBSimulatorDLLBackend(InputBackend):
    def __init__(self, driver: str = "AnyDriver"):
        self._driver = driver
        self._debug = str(os.getenv('WINDOWS_MCP_INPUT_DEBUG', '0')).lower() in ('1','true','yes','on')
        dll_path = _ib_dll_path()
        self._dll_path = str(dll_path) if dll_path else ""
        self._ready = False
        self._err = ""
        try:
            if not self._dll_path or not Path(self._dll_path).exists():
                self._err = f"dll not found at {self._dll_path}"
                return
            self._dll = ctypes.WinDLL(self._dll_path)
            self._dll.IbSendInit.argtypes = [ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p]
            self._dll.IbSendInit.restype = ctypes.c_uint32
            self._dll.IbSendDestroy.argtypes = []
            self._dll.IbSendDestroy.restype = None
            self._dll.IbSendInput.argtypes = [ctypes.c_uint, ctypes.c_void_p, ctypes.c_int]
            self._dll.IbSendInput.restype = ctypes.c_uint
            self._dll.IbSendMouseMove.argtypes = [ctypes.c_uint32, ctypes.c_uint32, ctypes.c_uint32]
            self._dll.IbSendMouseMove.restype = ctypes.c_bool
            self._dll.IbSendMouseClick.argtypes = [ctypes.c_uint32]
            self._dll.IbSendMouseClick.restype = ctypes.c_bool
            self._dll.IbSendMouseWheel.argtypes = [ctypes.c_int32]
            self._dll.IbSendMouseWheel.restype = ctypes.c_bool
            self._dll.IbSendKeybdDown.argtypes = [ctypes.c_uint16]
            self._dll.IbSendKeybdDown.restype = ctypes.c_bool
            self._dll.IbSendKeybdUp.argtypes = [ctypes.c_uint16]
            self._dll.IbSendKeybdUp.restype = ctypes.c_bool
            send_type = {
                "AnyDriver": 0,
                "SendInput": 1,
                "Logitech": 2,
                "Razer": 3,
                "DD": 4,
                "MouClassInputInjection": 5,
                "LogitechGHubNew": 6,
            }.get(self._driver, 0)
            rc = self._dll.IbSendInit(send_type, 0, None)
            if rc != 0:
                self._err = f"IbSendInit error={rc} (driver={self._driver})"
                return
            self._user32 = ctypes.WinDLL('user32', use_last_error=True)
            self._user32.VkKeyScanW.argtypes = [wintypes.WCHAR]
            self._user32.VkKeyScanW.restype = ctypes.c_short
            self._user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
            self._user32.SetCursorPos.restype = ctypes.c_bool
            self._ready = True
        except Exception as e:
            self._err = str(e)
            self._ready = False

    def info(self) -> BackendInfo:
        return BackendInfo("IBSimulatorDLL", self._ready, f"dll={self._dll_path}, driver={self._driver}, err={self._err}")

    # Mouse
    def move(self, x: int, y: int):
        if not self._ready:
            return
        xi, yi = int(x), int(y)
        self._user32.SetCursorPos(xi, yi)
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        pt = POINT()
        self._user32.GetCursorPos(ctypes.byref(pt))
        dx = xi - int(pt.x)
        dy = yi - int(pt.y)
        if dx or dy:
            self._dll.IbSendMouseMove(ctypes.c_uint32(dx & 0xFFFFFFFF).value,
                                      ctypes.c_uint32(dy & 0xFFFFFFFF).value,
                                      1)

    def click(self, x: int, y: int, button: str = "left", clicks: int = 1):
        if not self._ready:
            return
        self.move(x, y)
        btn_map = { 'left': 0x06, 'right': 0x18, 'middle': 0x60 }
        code = btn_map.get(button.lower(), 0x06)
        for _ in range(max(1, int(clicks))):
            self._dll.IbSendMouseClick(code)

    def drag(self, x1: int, y1: int, x2: int, y2: int):
        if not self._ready:
            return
        self.move(x1, y1)
        self._dll.IbSendMouseClick(0x02)  # LeftDown
        self.move(x2, y2)
        self._dll.IbSendMouseClick(0x04)  # LeftUp

    def _vk_for_key(self, key: str) -> int | None:
        if not isinstance(key, str) or not key:
            return None
        k = key.strip().lower()
        try:
            if k.startswith('vk'):
                return int(k[2:].lstrip('_'), 16)
            if k.startswith('0x'):
                return int(k, 16)
        except Exception:
            pass
        if k.startswith('f') and k[1:].isdigit():
            n = int(k[1:])
            if 1 <= n <= 24:
                return 0x70 + (n - 1)
        special = {
            'enter': 0x0D, 'return': 0x0D, 'backspace': 0x08, 'tab': 0x09,
            'esc': 0x1B, 'escape': 0x1B, 'space': 0x20,
            'capslock': 0x14, 'caps': 0x14, 'numlock': 0x90, 'scrolllock': 0x91,
            'pause': 0x13, 'break': 0x13, 'printscreen': 0x2C, 'prtsc': 0x2C, 'prtscr': 0x2C,
            'insert': 0x2D, 'ins': 0x2D, 'delete': 0x2E, 'del': 0x2E,
            'home': 0x24, 'end': 0x23, 'pageup': 0x21, 'pgup': 0x21, 'pagedown': 0x22, 'pgdn': 0x22,
            'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
            'shift': 0x10, 'lshift': 0xA0, 'rshift': 0xA1,
            'ctrl': 0x11, 'control': 0x11, 'lctrl': 0xA2, 'rctrl': 0xA3,
            'alt': 0x12, 'lalt': 0xA4, 'ralt': 0xA5,
            'win': 0x5B, 'lwin': 0x5B, 'rwin': 0x5C, 'apps': 0x5D, 'menu': 0x5D,
            'numpad0': 0x60, 'num0': 0x60, 'kp0': 0x60,
            'numpad1': 0x61, 'num1': 0x61, 'kp1': 0x61,
            'numpad2': 0x62, 'num2': 0x62, 'kp2': 0x62,
            'numpad3': 0x63, 'num3': 0x63, 'kp3': 0x63,
            'numpad4': 0x64, 'num4': 0x64, 'kp4': 0x64,
            'numpad5': 0x65, 'num5': 0x65, 'kp5': 0x65,
            'numpad6': 0x66, 'num6': 0x66, 'kp6': 0x66,
            'numpad7': 0x67, 'num7': 0x67, 'kp7': 0x67,
            'numpad8': 0x68, 'num8': 0x68, 'kp8': 0x68,
            'numpad9': 0x69, 'num9': 0x69, 'kp9': 0x69,
            'numpad*': 0x6A, 'multiply': 0x6A, 'kp_multiply': 0x6A,
            'numpad+': 0x6B, 'add': 0x6B, 'kp_add': 0x6B,
            'numpad-': 0x6D, 'subtract': 0x6D, 'kp_subtract': 0x6D,
            'numpad.': 0x6E, 'decimal': 0x6E, 'kp_decimal': 0x6E,
            'numpad/': 0x6F, 'divide': 0x6F, 'kp_divide': 0x6F,
            'numpadenter': 0x0D,
            'semicolon': 0xBA, ';': 0xBA, 'oem_1': 0xBA,
            'equals': 0xBB, '=': 0xBB, 'oem_plus': 0xBB,
            'comma': 0xBC, ',': 0xBC, 'oem_comma': 0xBC,
            'minus': 0xBD, '-': 0xBD, 'oem_minus': 0xBD,
            'period': 0xBE, 'dot': 0xBE, '.': 0xBE, 'oem_period': 0xBE,
            'slash': 0xBF, '/': 0xBF, 'oem_2': 0xBF,
            'grave': 0xC0, '`': 0xC0, 'backquote': 0xC0, 'oem_3': 0xC0,
            'leftbracket': 0xDB, '[': 0xDB, 'oem_4': 0xDB,
            'backslash': 0xDC, '\\': 0xDC, 'oem_5': 0xDC, 'pipe': 0xDC,
            'rightbracket': 0xDD, ']': 0xDD, 'oem_6': 0xDD,
            'apostrophe': 0xDE, "'": 0xDE, 'quote': 0xDE, 'oem_7': 0xDE,
            'oem_8': 0xDF,
        }
        if k in special:
            return special[k]
        if len(k) == 1:
            ch = k
            if 'a' <= ch <= 'z':
                return ord(ch.upper())
            if '0' <= ch <= '9':
                return ord(ch)
        return None

    def send_text(self, text: str):
        if not self._ready or not text:
            return
        import time

        # Increased delay to ensure proper character ordering
        # 20ms is more reliable for preventing race conditions
        char_delay = 0.020  # 20ms delay between characters
        key_delay = 0.003   # 3ms between key down and up

        if self._debug:
            logger.info(f"[send_text] Starting to send {len(text)} characters: {repr(text)}")

        # Use VK mode exclusively for reliability
        # Unicode batch mode is disabled due to race condition issues
        SHIFT = 0x10
        CTRL = 0x11
        ALT = 0x12

        for idx, ch in enumerate(text):
            if self._debug:
                logger.info(f"[send_text] Character {idx}: {repr(ch)}")

            vkshort = self._user32.VkKeyScanW(ch)
            if vkshort == -1:
                if self._debug:
                    logger.warning(f"[send_text] VkKeyScanW failed for character: {repr(ch)}")
                continue

            vk = vkshort & 0xFF
            mods = (vkshort >> 8) & 0xFF

            try:
                # Press modifiers
                if mods & 0x01:
                    self._dll.IbSendKeybdDown(SHIFT)
                    time.sleep(0.001)
                if mods & 0x02:
                    self._dll.IbSendKeybdDown(CTRL)
                    time.sleep(0.001)
                if mods & 0x04:
                    self._dll.IbSendKeybdDown(ALT)
                    time.sleep(0.001)

                # Press and release the key
                self._dll.IbSendKeybdDown(vk)
                time.sleep(key_delay)
                self._dll.IbSendKeybdUp(vk)

                # Release modifiers in reverse order
                if mods & 0x04:
                    time.sleep(0.001)
                    self._dll.IbSendKeybdUp(ALT)
                if mods & 0x02:
                    time.sleep(0.001)
                    self._dll.IbSendKeybdUp(CTRL)
                if mods & 0x01:
                    time.sleep(0.001)
                    self._dll.IbSendKeybdUp(SHIFT)

                # Wait before next character
                time.sleep(char_delay)

                if self._debug:
                    logger.info(f"[send_text] Character {idx} sent successfully")

            except Exception as e:
                if self._debug:
                    logger.error(f"[send_text] Error sending character {idx} ({repr(ch)}): {e}")
                continue

        if self._debug:
            logger.info(f"[send_text] Completed sending all characters")

    def hotkey(self, combo: str):
        if not self._ready:
            return
        parts = [p.strip().lower() for p in combo.split('+') if p.strip()]
        mods = set(); key = None
        for p in parts:
            if p in ("ctrl", "control"): mods.add("ctrl"); continue
            if p in ("alt",): mods.add("alt"); continue
            if p in ("shift",): mods.add("shift"); continue
            if p in ("win", "lwin", "rwin"): mods.add("win"); continue
            key = p
        if "win" in mods: self._dll.IbSendKeybdDown(0x5B)
        if "ctrl" in mods: self._dll.IbSendKeybdDown(0x11)
        if "alt" in mods: self._dll.IbSendKeybdDown(0x12)
        if "shift" in mods: self._dll.IbSendKeybdDown(0x10)
        if key:
            vk = self._vk_for_key(key)
            if vk is None and len(key) == 1:
                vkshort = self._user32.VkKeyScanW(key)
                vk = vkshort & 0xFF if vkshort != -1 else None
            if vk is not None:
                self._dll.IbSendKeybdDown(vk)
                self._dll.IbSendKeybdUp(vk)
        if "shift" in mods: self._dll.IbSendKeybdUp(0x10)
        if "alt" in mods: self._dll.IbSendKeybdUp(0x12)
        if "ctrl" in mods: self._dll.IbSendKeybdUp(0x11)
        if "win" in mods: self._dll.IbSendKeybdUp(0x5B)

    def key_down(self, key: str):
        if not self._ready:
            return
        vk = self._vk_for_key(key)
        if vk is None and len(key) == 1:
            vkshort = self._user32.VkKeyScanW(key)
            vk = vkshort & 0xFF if vkshort != -1 else None
        if vk is None:
            return
        self._dll.IbSendKeybdDown(int(vk))

    def key_up(self, key: str):
        if not self._ready:
            return
        vk = self._vk_for_key(key)
        if vk is None and len(key) == 1:
            vkshort = self._user32.VkKeyScanW(key)
            vk = vkshort & 0xFF if vkshort != -1 else None
        if vk is None:
            return
        self._dll.IbSendKeybdUp(int(vk))

    def scroll(self, wheel_times: int, type: str = "vertical", direction: str = "down"):
        if not self._ready:
            return
        t = (type or "vertical").lower(); d = (direction or "down").lower()
        WHEEL_DELTA = 120
        n = max(1, int(wheel_times))
        if t == "vertical":
            delta = WHEEL_DELTA * n
            if d == "down": delta = -delta
            try:
                self._dll.IbSendMouseWheel(int(delta))
            except Exception:
                pass
        elif t == "horizontal":
            try:
                user32 = ctypes.windll.user32
                MOUSEEVENTF_HWHEEL = 0x01000
                delta = WHEEL_DELTA * n
                if d == "left": delta = -delta
                elif d == "right": delta = +delta
                else: return
                user32.mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, int(delta), 0)
            except Exception:
                pass
