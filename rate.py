from __future__ import annotations
from dataclasses import dataclass
from time import monotonic, sleep


@dataclass
class RateConfig:
    mouse_move_hz: float = 120.0
    mouse_max_delta: int = 60
    mouse_smooth: float = 0.0
    clicks_per_sec: float = 8.0
    keys_per_sec: float = 12.0


class RateLimiter:
    def __init__(self, cfg: RateConfig | None = None):
        self.cfg = cfg or RateConfig()
        self._last_click = 0.0
        self._last_key = 0.0
        self._last_move = 0.0

    def update_config(self, cfg: RateConfig):
        self.cfg = cfg

    def time_until_click(self) -> float:
        min_dt = 1.0 / max(self.cfg.clicks_per_sec, 0.001)
        dt = monotonic() - self._last_click
        return max(0.0, min_dt - dt)

    def mark_click(self):
        self._last_click = monotonic()

    def time_until_key(self) -> float:
        min_dt = 1.0 / max(self.cfg.keys_per_sec, 0.001)
        dt = monotonic() - self._last_key
        return max(0.0, min_dt - dt)

    def mark_key(self):
        self._last_key = monotonic()

    def time_until_move(self) -> float:
        min_dt = 1.0 / max(self.cfg.mouse_move_hz, 0.001)
        dt = monotonic() - self._last_move
        return max(0.0, min_dt - dt)

    def mark_move(self):
        self._last_move = monotonic()

    def filter_target(self, cur: tuple[int, int], target: tuple[int, int]) -> tuple[int, int]:
        cx, cy = cur
        tx, ty = target
        dx, dy = tx - cx, ty - cy
        md = self.cfg.mouse_max_delta
        if md > 0:
            if abs(dx) > md: dx = md if dx > 0 else -md
            if abs(dy) > md: dy = md if dy > 0 else -md
        s = self.cfg.mouse_smooth
        if s > 0:
            dx = int(dx * (1.0 - s))
            dy = int(dy * (1.0 - s))
        return cx + dx, cy + dy

    def sleep_until_ready(self, kind: str):
        if kind == 'click':
            t = self.time_until_click();
            if t > 0: sleep(t)
            self.mark_click()
        elif kind == 'key':
            t = self.time_until_key();
            if t > 0: sleep(t)
            self.mark_key()
        elif kind == 'move':
            t = self.time_until_move();
            if t > 0: sleep(t)
            self.mark_move()

