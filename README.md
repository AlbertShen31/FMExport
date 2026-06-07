# fm26-export-tool

A **macOS-first** Python CLI that converts Football Manager 2026 **Web Page** exports into CSV, Excel, JSON, Parquet, and SQLite.

This tool takes a safe, reliable path: you export what you see in-game, and this CLI parses the saved HTML.

## How to export from FM26 on Mac

1. Open a player list view (squad, search results, scouting shortlist, etc.).
2. Click the **first row** to focus the list.
3. Press **⌘+A** to select all visible rows.
4. Press **⌘+P** to open the print dialog.
5. Choose **Web Page** as the output format.
6. Save the `.html` file (e.g. to `~/Downloads` or `~/Documents`).

Then convert it:

```bash
fm26-export export-csv ~/Downloads/players.html
fm26-export export-xlsx ~/Downloads/players.html
fm26-export inspect-columns ~/Downloads/players.html
```

## Why Web Page export

- **No game modification** — FM files are never touched.
- **No runtime injection** — nothing hooks into the game process.
- **No memory patching or save-file parsing** — only reads HTML you explicitly exported.
- **Matches what you see** — exports reflect your active FM view and column layout.
- **Easy to audit** — HTML is plain text you can open and inspect.

## Installation

Requires **Python 3.11+**.

```bash
# From the project root
pip install -e .

# Optional Parquet support
pip install -e ".[parquet]"

# Development / tests
pip install -e ".[dev]"
```

## CLI commands

| Command | Description |
|---------|-------------|
| `scan` | Detect likely FM26 folders on macOS |
| `parse-html <file>` | Parse one HTML export and print a JSON summary |
| `export-csv <file>` | Write UTF-8 CSV |
| `export-xlsx <file>` | Write Excel with autofilter and frozen header |
| `export-json <file>` | Write JSON records |
| `export-parquet <file>` | Write Parquet (requires `pyarrow`) |
| `export-sqlite <file>` | Write SQLite database |
| `inspect-columns <file>` | Show detected columns and sample values |
| `watch-folder [dir]` | Auto-convert new HTML files as they appear |

### Examples

```bash
# Find FM26-related folders
fm26-export scan

# Inspect columns before exporting
fm26-export inspect-columns ~/Downloads/squad.html

# Export to multiple formats
fm26-export export-csv ~/Downloads/squad.html -o ~/Desktop/squad.csv
fm26-export export-xlsx ~/Downloads/squad.html -o ~/Desktop/squad.xlsx
fm26-export export-json ~/Downloads/squad.html

# Watch Downloads for new exports
fm26-export watch-folder ~/Downloads --formats csv,xlsx,json
```

## Parser behavior

- Uses **BeautifulSoup** + **pandas.read_html** to extract tables.
- Picks the table most likely to be a player list (Name, Age, Position, etc.).
- **Normalizes** column names to `snake_case` while preserving originals in metadata.
- Handles **duplicate column names** (`value`, `value_2`, …).
- Normalizes **currency** (`£12.5M`), **wages** (`£50K p/w`), **dates**, **star ratings** (★★★☆☆), and empty values (`-`, `N/A`).

Use `--no-normalize` to keep raw cell text.

## Project layout

```
fm26_export_tool/
  __init__.py
  cli.py          # Click CLI
  parser.py       # HTML table extraction
  normalize.py    # Column/value normalization
  exporters.py    # CSV, XLSX, JSON, Parquet, SQLite
  paths.py        # macOS FM26 folder detection
  watcher.py      # Folder watch auto-convert
tests/
  test_parser.py
  test_normalize.py
```

## Screenshot OCR export (alternative)

For screenshot-based extraction without HTML exports, see [`fm26-screenshot-exporter/`](fm26-screenshot-exporter/README.md) (`fm26-ocr` CLI).

This project does **not** implement memory patching, save-file parsing, DRM bypassing, or any modification of FM game files.

## Running tests

```bash
pytest
```

## License

MIT
