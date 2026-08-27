"""
progress.py
-------------
A minimal terminal spinner for showing "still working" feedback during
blocking operations (like reading a large CSV). Standard-library only —
no extra dependency needed just for this.

Usage:
    with Spinner("Loading conversations.csv"):
        df = pd.read_csv(...)
"""

import sys
import threading
import itertools
import time


class Spinner:
    """Context manager that animates a spinner + label until the block exits."""

    FRAMES = ["|", "/", "-", "\\"]

    def __init__(self, label: str):
        self.label = label
        self._stop_event = threading.Event()
        self._thread = threading.Thread(target=self._spin, daemon=True)

    def _spin(self):
        for frame in itertools.cycle(self.FRAMES):
            if self._stop_event.is_set():
                break
            sys.stdout.write(f"\r{self.label}... {frame}")
            sys.stdout.flush()
            time.sleep(0.1)

    def __enter__(self):
        self._thread.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._stop_event.set()
        self._thread.join()
        # Clear the spinner line and print a static "done" line in its place.
        status = "failed" if exc_type else "done"
        sys.stdout.write(f"\r{self.label}... {status}\n")
        sys.stdout.flush()
        return False  # don't suppress exceptions