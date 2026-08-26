"""Pacing and interruption: the adaptive key throttle and the F12 abort watch.

Both are loops over measured time rather than guessed time -- the throttle
derives its delay from observed latency and may only ever slow down, and the
abort watch polls a key rather than installing a hook.

No win32api here: AbortWatch polls _user32.GetAsyncKeyState, which sees the key
regardless of which window has focus and needs no extra dependency.
"""

from __future__ import annotations

import threading
import time

from tme import config
from tme.win.errors import Aborted
from tme.win.native import _user32

# --- BEGIN MOVED FROM tmeio.py ---
# --------------------------------------------------------------------------
# Adaptive throttle
# --------------------------------------------------------------------------


class Throttle:
    """Key delay and read timeout derived from measured TME latency.

    Fast while the app is responsive, automatically slower when it stutters --
    which is the whole point: a fixed delay is either too slow all run or too
    fast exactly when the app is struggling.
    """

    def __init__(self) -> None:
        self.baseline_s = 0.05
        self.ewma_s = self.baseline_s
        self.key_delay_s = config.KEY_DELAY_START_S
        self.read_timeout_s = config.READ_TIMEOUT_START_S
        self.consecutive_timeouts = 0
        self.penalties = 0
        self.max_latency_s = 0.0

    @staticmethod
    def _clamp(value: float, low: float, high: float) -> float:
        return max(low, min(high, value))

    def observe(self, latency_s: float) -> None:
        """Fold a successful read's latency into the running estimate."""
        self.consecutive_timeouts = 0
        self.max_latency_s = max(self.max_latency_s, latency_s)
        self.ewma_s = 0.75 * self.ewma_s + 0.25 * latency_s
        ratio = self.ewma_s / self.baseline_s
        self.key_delay_s = self._clamp(
            config.KEY_DELAY_START_S * ratio,
            config.KEY_DELAY_MIN_S,
            config.KEY_DELAY_MAX_S,
        )
        self.read_timeout_s = self._clamp(
            self.ewma_s * 6.0,
            config.READ_TIMEOUT_MIN_S,
            config.READ_TIMEOUT_MAX_S,
        )

    def penalize(self) -> None:
        """A read timed out: back off hard before retrying."""
        self.consecutive_timeouts += 1
        self.penalties += 1
        self.key_delay_s = self._clamp(
            self.key_delay_s * 2.0, config.KEY_DELAY_MIN_S, config.KEY_DELAY_MAX_S
        )
        self.read_timeout_s = self._clamp(
            self.read_timeout_s * 2.0,
            config.READ_TIMEOUT_MIN_S,
            config.READ_TIMEOUT_MAX_S,
        )

    def summary(self) -> str:
        return (
            f"latency avg {self.ewma_s * 1000:.0f} ms, peak "
            f"{self.max_latency_s * 1000:.0f} ms, key delay "
            f"{self.key_delay_s * 1000:.0f} ms, {self.penalties} backoff(s)"
        )


# --------------------------------------------------------------------------
# Abort watchdog
# --------------------------------------------------------------------------


class AbortWatch:
    """Polls the abort hotkey in a daemon thread.

    ``GetAsyncKeyState`` sees the key regardless of which window has focus, so
    the hotkey still works while TME is in front and needs no extra dependency.
    """

    def __init__(self) -> None:
        self.event = threading.Event()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        def loop() -> None:
            while not self._stop.is_set():
                if _user32.GetAsyncKeyState(config.ABORT_VK) & 0x8000:
                    self.event.set()
                    return
                time.sleep(0.03)

        self._thread = threading.Thread(target=loop, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()

    def check(self) -> None:
        if self.event.is_set():
            raise Aborted(f"aborted by {config.ABORT_KEY_NAME}")


