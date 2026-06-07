# FM26 Screenshot Exporter

macOS-first Football Manager 2026 screenshot OCR/stat parser. Converts screenshots of FM player/stat tables into structured data: **CSV**, **Excel**, **JSON**, and optionally **Parquet**.

**No game modification.** No BepInEx. No memory patching. Screenshots only.

## Requirements

- macOS (primary target)
- Python 3.11+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) via Homebrew:

```bash
brew install tesseract
```

## Install

```bash
cd fm26-screenshot-exporter
pip install -e ".[dev]"
```

Optional extras:

```bash
pip install -e ".[parquet]"      # Parquet export
pip install -e ".[streamlit]"      # Future UI experiments
pip install -e ".[easyocr]"        # Alternative OCR engine (not used by default)
```

## Workflow

1. In FM26, set **windowed mode** or a consistent resolution.
2. Use the same skin/theme and zoom level every time.
3. Open a player search or stats table and show all desired columns.
4. Take a screenshot:
   - **Cmd + Shift + 5** → capture selected window or screen
   - Save to a known folder (e.g. `~/Pictures/FM26/`)
5. Create or adjust a profile YAML in `profiles/` (start from `player_search_default.yaml`).
6. Verify your environment:

```bash
fm26-ocr scan-env
```

7. Generate a debug overlay and tune crop/column boundaries:

```bash
fm26-ocr debug screenshot.png --profile player_search_default --out debug.png
```

8. Parse and export:

```bash
fm26-ocr parse screenshot.png --profile player_search_default --out players.csv
fm26-ocr parse screenshot.png --profile player_search_default --out players.xlsx
fm26-ocr parse screenshot.png --profile player_search_default --out players.json
```

9. Open `debug.png` and adjust `manual_crop` and `fixed_columns` in your profile until columns align.

## CLI Commands

| Command | Description |
|---------|-------------|
| `scan-env` | Check Tesseract + Python package versions |
| `preprocess IMAGE --out OUT` | Crop, upscale, denoise, sharpen, threshold |
| `ocr IMAGE --out raw_ocr.json` | Word-level OCR with bounding boxes |
| `detect-table IMAGE --out table_regions.json` | Detect table/header/row/column regions |
| `parse IMAGE --profile NAME --out OUT` | Full pipeline → CSV/XLSX/JSON/Parquet |
| `parse-folder DIR --profile NAME --out-dir OUT` | Batch process PNG/JPG screenshots |
| `watch-folder DIR --profile NAME --out-dir OUT` | Auto-parse new screenshots |
| `debug IMAGE --profile NAME --out debug.png` | Overlay columns, rows, OCR boxes |

### Examples

```bash
# Environment check
fm26-ocr scan-env

# Step-by-step debugging
fm26-ocr preprocess screenshot.png --profile player_search_default --out preprocessed.png
fm26-ocr ocr screenshot.png --profile player_search_default --out raw_ocr.json
fm26-ocr detect-table screenshot.png --profile player_search_default --out table_regions.json

# Full parse
fm26-ocr parse screenshot.png --profile player_search_default --out players.csv

# Batch folder
fm26-ocr parse-folder ~/Pictures/FM26 --profile player_search_default --out-dir ./exports --format xlsx

# Watch folder (Ctrl+C to stop)
fm26-ocr watch-folder ~/Pictures/FM26 --profile player_search_default --out-dir ./exports
```

## Profiles

Profiles are YAML files in `profiles/`. They define crop regions, fixed column x-boundaries, normalizers, and validation thresholds.

Key fields:

- `manual_crop` — table region `{x, y, width, height}`
- `fixed_columns` — list of `{name, x_start, x_end}` (recommended for MVP)
- `header_y_range` — optional header row y-range (relative to crop)
- `row_y_tolerance` — pixel tolerance for row clustering
- `min_confidence` — drop OCR tokens below this threshold
- `value_normalizers` — per-column parsers (`age`, `currency`, `wage`, etc.)

See `profiles/player_search_default.yaml` and `examples/sample_config.yaml`.

## Recommended FM Screenshot Settings

- Use **windowed mode** (not fullscreen with OS scaling changes).
- Avoid display scaling changes between sessions.
- Use the **maximum readable font size** in your skin.
- Keep columns **fixed** between screenshots — same view, same sort order.
- Avoid transparent overlays or tooltips over the table.
- Sort consistently (e.g. always by name or value the same way).
- Screenshot **one full table page** at a time.
- For multiple pages, use `parse-folder` and combine outputs in Excel or pandas.

## Accuracy Strategy

1. **Start with fixed crop + fixed column boundaries** — FM screenshots are visually consistent; this beats generic table detection.
2. Use the **debug overlay** (`fm26-ocr debug`) after every profile change.
3. **Tune the profile once** for your resolution/skin/zoom, then reuse it.
4. Store **raw OCR JSON** (`fm26-ocr ocr`) for auditing misparsed cells.
5. Review **validation warnings** printed after each parse (low confidence, empty cells, duplicates).
6. Re-export to **XLSX** when you want both parsed and raw OCR sheets side by side.

## Excel Output

`.xlsx` exports include:

- **Parsed Data** — normalized values, frozen header, autofilter, auto column widths
- **Raw OCR** — reconstructed cell text before normalization
- **OCR Tokens** — full token list with coordinates and confidence
- **Metadata** — source path, profile, row counts, validation summary

## Project Structure

```
fm26_screenshot_exporter/
  cli.py                      # Typer CLI
  config.py                   # Pydantic models
  image_preprocess.py         # Crop, upscale, denoise, sharpen
  table_detection.py          # OpenCV line detection + manual crop
  ocr.py                      # pytesseract wrapper
  row_column_reconstruction.py  # Row/column assignment
  normalize.py                # Value parsers
  exporters.py                # CSV, XLSX, JSON, Parquet
  profiles.py                 # YAML profile loader
  validators.py               # Parse quality warnings
  watcher.py                  # Folder watcher
  debug_render.py             # Debug overlay renderer
  pipeline.py                 # End-to-end parse orchestration
profiles/
  player_search_default.yaml
tests/
examples/
```

## Testing

```bash
cd fm26-screenshot-exporter
pytest
```

## Limitations

- OCR accuracy depends on screenshot quality, skin contrast, and consistent layout.
- Automatic table detection is a fallback; **fixed column profiles** are the recommended path.
- Does not read game memory, modify FM files, or require BepInEx.
- Column x-coordinates are resolution-dependent — retune if you change window size.

## License

MIT
