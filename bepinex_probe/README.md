# FM26 BepInEx Export Probe (macOS, experimental)

Research-oriented, **staged** probe for testing whether BepInEx 6 can load on FM26 macOS (Unity IL2CPP) and whether a minimal plugin can observe in-game UI structure.

This module does **not** implement real player export yet. It only validates prerequisites before any UI reverse-engineering work begins.

## Safety constraints

- No DRM, anti-cheat, signature, or online-protection bypassing
- No direct memory patching
- No save-file parsing or modification
- No overwriting game files without an explicit backup
- Goal: test plugin load + write reversible diagnostics only

If BepInEx cannot boot on your Mac build, use the main project's **manual Web Page export** workflow (`fm26-export` CLI) instead.

---

## Staged workflow

### Stage 1 — Detect FM26 Unity backend

```bash
python3 bepinex_probe/scripts/detect_fm26_backend.py \
  --app-path "/Users/$USER/Library/Application Support/Steam/steamapps/common/Football Manager 26"
```

On Steam macOS installs the bundle is often `fm.app` (not `Football Manager 26.app`). The scripts resolve that automatically when you pass the install directory.

Inspects:

- `Contents/Resources/Data/Managed`
- `Contents/Resources/Data/il2cpp_data`
- `Contents/Frameworks`
- `Contents/MacOS`

Prints:

- Unity **Mono** vs **IL2CPP** (or `unknown`)
- Likely executable path
- Warnings if backend cannot be determined

**Do not install BepInEx** until the backend is confirmed.

---

### Stage 2 — Detect BepInEx install + logs

Only after you have **manually** installed a compatible BepInEx 6 IL2CPP macOS build (if one exists for your setup):

```bash
python3 bepinex_probe/scripts/detect_bepinex_install.py \
  --app-path "/path/to/Football Manager 26.app"
```

Looks for:

- `BepInEx/`
- `BepInEx/plugins/`
- `BepInEx/LogOutput.log`
- `run_bepinex.sh`
- `doorstop_config.ini`

Prints:

- Whether BepInEx appears installed
- Last 100 lines of `LogOutput.log`
- Marker hits: `Chainloader startup complete`, `Loading plugin`, `FM26ExportProbe`, `error`, `exception`

**Launch FM26 once** through BepInEx before expecting logs.

---

### Stage 3 — Verify plugin skeleton

```bash
python3 bepinex_probe/scripts/create_plugin_skeleton.py
```

Confirms source files exist under `src/FM26ExportProbe/`.

---

### Stage 4 — Build and install probe plugin

**Backup first.** Copy your FM26.app or game folder before copying any DLLs.

```bash
# Point to BepInEx core DLLs from your FM26 install
export BEPINEX_LIBS="/path/to/Football Manager 26.app/Contents/MacOS/BepInEx/core"

python3 bepinex_probe/scripts/package_plugin.py
```

Output:

```
bepinex_probe/dist/FM26ExportProbe/
  FM26ExportProbe.dll
  INSTALL.txt
```

Copy into the game:

```
<Football Manager 26.app>/Contents/MacOS/BepInEx/plugins/FM26ExportProbe/FM26ExportProbe.dll
```

---

### Stage 5 — Launch FM26 and confirm probe load

1. Launch FM26 (via BepInEx / `run_bepinex.sh` if required).
2. Re-run Stage 2 detection and check logs.
3. Confirm file created:

```
~/Documents/FM26Exports/probe_loaded.txt
```

4. In-game, press **F8**.
5. Inspect:

```
~/Documents/FM26Exports/fm26_probe_diagnostic.txt
```

Diagnostics include:

- Loaded assemblies (sample)
- Unity scene names
- Root GameObject names
- UI component type counts
- Keyword hits for likely table/grid/list nodes

---

### Stage 6 — Only after probe success: UI table research

`Exporter.cs` currently writes diagnostics only.

`UiScanner.cs` performs a first-pass scene tree walk and looks for generic names:

`table`, `grid`, `row`, `cell`, `list`, `squad`, `player`, `search`, `browser`

Do **not** assume FM26 class names yet. Use diagnostics to identify candidate UI nodes before building a real exporter.

`CsvWriter.cs` is a generic helper for later CSV output.

---

## Plugin details

| Field | Value |
|-------|-------|
| GUID | `com.local.fm26.exportprobe` |
| Name | FM26 Export Probe |
| Version | `0.1.0` |

**Awake:**

- Logs `FM26 Export Probe loaded`
- Creates `~/Documents/FM26Exports/`
- Writes `probe_loaded.txt` with UTC timestamp

**Update (F8):**

- Writes `fm26_probe_diagnostic.txt` via `Exporter.WriteDiagnostics()`

---

## Success criteria before real export work

All of the following must pass:

1. BepInEx creates `LogOutput.log`
2. Log contains `Chainloader startup complete`
3. Plugin `Awake` runs (log message + `probe_loaded.txt` exists)
4. **F8** writes `fm26_probe_diagnostic.txt`
5. Diagnostics include scene names, root GameObjects, and UI component counts

If any step fails, **stop** and use manual HTML export (`fm26-export` in the main repo).

---

## Stop conditions

Fall back to **manual Web Page export** or screenshot OCR if:

- BepInEx does not boot on macOS
- `LogOutput.log` is never created
- Plugin never loads (no `probe_loaded.txt`, no `FM26ExportProbe` in logs)
- FM26 crashes on launch after plugin install
- Logs contain unrecoverable errors/exceptions during chainload

These are expected outcomes on macOS IL2CPP — the HTML export pipeline remains the supported path.

---

## Reversibility

- Plugin DLL lives only in `BepInEx/plugins/FM26ExportProbe/` — remove that folder to uninstall
- Probe writes only to `~/Documents/FM26Exports/` — safe to delete
- Detection scripts are read-only
- Always keep a backup of `Football Manager 26.app` before installing BepInEx

---

## Build requirements

- .NET 6 SDK (`dotnet --version`)
- BepInEx 6 IL2CPP core DLLs from your FM26 install (`BEPINEX_LIBS`)
- FM26 macOS app bundle for runtime testing

---

## Related

Main supported export path: [../README.md](../README.md) — FM26 Web Page HTML → CSV/Excel/JSON.
