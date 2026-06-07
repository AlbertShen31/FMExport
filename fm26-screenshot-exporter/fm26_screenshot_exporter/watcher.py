"""Watch a folder for new screenshots and auto-parse them."""

from __future__ import annotations

import time
from pathlib import Path
from typing import Callable

from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp"}


def is_image_file(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_SUFFIXES


class ScreenshotHandler(FileSystemEventHandler):
    def __init__(
        self,
        on_image: Callable[[Path], None],
        *,
        settle_seconds: float = 1.0,
    ) -> None:
        self.on_image = on_image
        self.settle_seconds = settle_seconds
        self._pending: dict[str, float] = {}

    def on_created(self, event: FileSystemEvent) -> None:
        if event.is_directory:
            return
        self._schedule(Path(event.src_path))

    def on_modified(self, event: FileSystemEvent) -> None:
        if event.is_directory:
            return
        self._schedule(Path(event.src_path))

    def _schedule(self, path: Path) -> None:
        if not is_image_file(path):
            return
        self._pending[str(path)] = time.time()

    def process_pending(self) -> None:
        now = time.time()
        ready = [
            Path(p)
            for p, ts in list(self._pending.items())
            if now - ts >= self.settle_seconds and Path(p).exists()
        ]
        for path in ready:
            self._pending.pop(str(path), None)
            self.on_image(path)


def watch_folder(
    input_dir: str | Path,
    on_image: Callable[[Path], None],
    *,
    settle_seconds: float = 1.0,
    poll_interval: float = 0.5,
) -> None:
    """Block and watch input_dir for new image files."""
    root = Path(input_dir)
    root.mkdir(parents=True, exist_ok=True)

    handler = ScreenshotHandler(on_image, settle_seconds=settle_seconds)
    observer = Observer()
    observer.schedule(handler, str(root), recursive=False)
    observer.start()

    try:
        while True:
            handler.process_pending()
            time.sleep(poll_interval)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
