"""Watch a folder for new FM HTML exports and auto-convert them."""

from __future__ import annotations

import time
from pathlib import Path

from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

from fm26_export_tool.exporters import export_all
from fm26_export_tool.parser import parse_html


class FMExportHandler(FileSystemEventHandler):
    """Process new HTML files dropped into a watch folder."""

    def __init__(
        self,
        *,
        formats: list[str],
        output_dir: Path | None = None,
        debounce_seconds: float = 1.0,
    ) -> None:
        super().__init__()
        self.formats = formats
        self.output_dir = output_dir
        self.debounce_seconds = debounce_seconds
        self._pending: dict[str, float] = {}
        self._processed: set[str] = set()

    def _should_handle(self, path: Path) -> bool:
        return path.suffix.lower() in {".html", ".htm"} and path.is_file()

    def _process(self, path: Path) -> None:
        key = str(path.resolve())
        if key in self._processed:
            return

        try:
            parsed = parse_html(path)
            results = export_all(parsed, self.output_dir, formats=self.formats)
            self._processed.add(key)
            print(f"[fm26-export] Converted {path.name}:")
            for fmt, out in results.items():
                print(f"  {fmt}: {out}")
        except Exception as exc:
            print(f"[fm26-export] Failed to process {path}: {exc}")

    def on_created(self, event: FileSystemEvent) -> None:
        if event.is_directory:
            return
        path = Path(event.src_path)
        if self._should_handle(path):
            self._pending[str(path)] = time.monotonic()

    def on_modified(self, event: FileSystemEvent) -> None:
        if event.is_directory:
            return
        path = Path(event.src_path)
        if self._should_handle(path):
            self._pending[str(path)] = time.monotonic()

    def drain_pending(self) -> None:
        now = time.monotonic()
        ready = [
            p
            for p, ts in list(self._pending.items())
            if now - ts >= self.debounce_seconds
        ]
        for path_str in ready:
            del self._pending[path_str]
            self._process(Path(path_str))


def watch_folder(
    folder: str | Path,
    *,
    formats: list[str] | None = None,
    output_dir: str | Path | None = None,
    poll_interval: float = 0.5,
) -> None:
    """Watch a folder and auto-convert new FM HTML exports."""
    watch_path = Path(folder).expanduser().resolve()
    if not watch_path.is_dir():
        raise NotADirectoryError(f"Watch folder does not exist: {watch_path}")

    selected_formats = formats or ["csv", "xlsx", "json"]
    out = Path(output_dir).expanduser().resolve() if output_dir else None
    if out:
        out.mkdir(parents=True, exist_ok=True)

    handler = FMExportHandler(formats=selected_formats, output_dir=out)
    observer = Observer()
    observer.schedule(handler, str(watch_path), recursive=False)
    observer.start()

    print(f"[fm26-export] Watching {watch_path} for HTML exports...")
    print(f"[fm26-export] Output formats: {', '.join(selected_formats)}")
    if out:
        print(f"[fm26-export] Output directory: {out}")
    print("[fm26-export] Press Ctrl+C to stop.")

    try:
        while True:
            handler.drain_pending()
            time.sleep(poll_interval)
    except KeyboardInterrupt:
        print("\n[fm26-export] Stopping watcher.")
    finally:
        observer.stop()
        observer.join()
