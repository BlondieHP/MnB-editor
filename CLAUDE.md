# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**魔球修改器 (MnBWarband Editor / "Mojo")** — a Visual Basic 6.0 editor for Mount & Blade: Warband mod files. It edits the game's TXT-format data files (items, troops, factions, parties, scenes, particle systems, sounds, meshes, tableau materials, map icons, triggers, strings, global variables) and their corresponding CSV localization files.

- **Language**: Visual Basic 6.0 (source files: `.bas` modules, `.frm` forms, `.cls` classes, `.ctl` user controls)
- **Entry point**: `Welcome.frm` (Startup form) → `frmMain.frm` (MDI parent)
- **Source encoding**: **GBK (cp936)** — All `.bas`, `.frm`, `.cls`, `.ctl`, and `.ini` files use the Windows Chinese ANSI code page. VB6 IDE requires this encoding and does NOT support UTF-8. Chinese comments will appear garbled in UTF-8-mode editors. To read a file with correct Chinese rendering: `iconv -f GBK -t UTF-8 <file>`
- **No build system, no tests**

## Building & Running

The only way to build is via the VB6 IDE (`VB6.EXE`). Open `MnBWarband_Editor.vbp` and press F5 (Run) or File → Make EXE. All compiler optimizations are disabled in the project settings (no bounds checking, no overflow checking, no aliasing).

The `Release/` directory contains a pre-built `MnBWarband_Editor.exe` (3.5MB, compiled 2013-06-28).

The `MnBWarband_Editor(Setup)/` directory contains installer packaging files.

## Architecture

### Data Flow

```
TXT files (troops.txt, item_kinds1.txt, etc.)
    │
    ▼ byte-level random-access read (GetWord / GetLine)
Global arrays (Trps(), itm(), Factions(), etc.)
    │
    ▼ Form controls bind directly to global arrays
User edits → global arrays updated in-place
    │
    ▼ byte-level write (PutWord / PutReturn)
Back to TXT files
```

```
CSV files (troops.csv, item_kinds.csv, etc.) in ModPath\languages\<Language>\
    │
    ▼ UEFLoadTextFile (UTF-8 aware, from mTextUTF.bas) → Split
Global struct fields (.csvName, .csvName_pl)
    │
    ▼ Displayed in ListViews alongside internal IDs
```

There is **no abstraction layer** — forms read/write global arrays directly. There is **no data validation** between parsing and display.

### Module Map

| Module | Lines | Role |
|--------|-------|------|
| `ModMain.bas` | 9,864 | **God module**: all Type definitions, all global variables, all file load/save, CSV load, utility functions, color conversion, entity enumeration |
| `ModOperation.bas` | 5,098 | Operation code constants (~1500 game scripting opcodes), operation-to-Python/CSV mapping, parameter type definitions |
| `ModPython.bas` | 2,236 | Python-code generation for triggers/operations, game flag/attribute/skill registry, trigger function presets |
| `ModCoder.bas` | 1,220 | Python/TXT/Pseudo code ↔ internal OpBlock encoding/decoding, parameter splitting/standardization |
| `ModApp.bas` | 633 | Application-level utilities (error logging, backup filename generation, random, rounding, window resize, tree helpers) |
| `MdlLanMgr.bas` | ~600 | Language/translation manager — loads `.lan.ini` files, provides UI string translations |
| `ModBackUp.bas` | 204 | Backup/restore of all mod data files (hardcoded file list, no transaction support) |
| `ModMemory.bas` | 209 | Low-level memory utilities via `CopyMemory` (HiByte, LoByte, MakeInteger, MakeLong, etc.) |
| `UInt64.bas` | ~700 | 64-bit integer type (`Integer64b`) — arithmetic, hex conversion, bit manipulation for game IDs |
| `mTextUTF.bas` | 458 | UTF-8/UTF-16 file I/O via Win32 API (third-party, by zyl910) |
| `Win32API.bas` | ~500 | Win32 API declarations |
| `ModINI.bas` | 51 | Loads module.ini resource declarations |
| `ModTreeView.bas` | ~1000 | TreeView population for troop upgrade trees |
| `ModHook.bas` | ~50 | Windows hook for IDE debugging |
| `clsOpBlock.cls` | 579 | Operation block editor bridge — TXT/PY/Pseudo three-way encoding sync |

### Key Forms

| Form | Size | Tag | Edits |
|------|------|-----|-------|
| `frmItems.frm` | 195KB | `edit_2` | Items (equipment, weapons, goods) |
| `frmTroops.frm` | 119KB | `edit_1` | Troops (NPCs, soldiers, heroes) |
| `frmParties.frm` | 114KB | `edit_3` | Parties (map parties) |
| `frmParty_Templates.frm` | 78KB | — | Party templates |
| `frmFactions.frm` | 39KB | — | Factions |
| `frmScenes.frm` | 66KB | — | Scenes |
| `frmMap_Icons.frm` | 42KB | — | Map icons |
| `frmPSys.frm` | 69KB | — | Particle systems |
| `frmMesh.frm` | 37KB | — | Meshes |
| `frmSounds.frm` | 38KB | — | Sounds |
| `frmSoundRess.frm` | 31KB | — | Sound resources |
| `frmTrigger.frm` | 45KB | — | Simple triggers |
| `frmTabMat.frm` | 43KB | — | Tableau materials |
| `frmStrTool.frm` | 19KB | — | String editor |
| `frmMain.frm` | 31KB | — | Main MDI container window |

Forms are identified by their `Tag` property (e.g., `edit_1` for troops) which maps to entries in `PublicEditors()` array loaded from language INI files.

### Custom Controls (in `Controls/`)

- **ListViewforMS** — Enhanced ListView for mod-structure display (with icon, ID, and name columns)
- **ListViewforPY** — ListView variant for Python code parameter display
- **MenuforMS** — Custom popup menu control
- **MyTab** — Custom tab strip
- **OpBlockEditor** — Three-tab operation block editor (Pseudo-code / Python / TXT views)
- **ComboforOp** — Dropdown with operation-specific parameter type filtering
- **ComboEx** — Extended combo box
- **RichforPY** — RichTextBox wrapper for Python syntax
- **RichforTXT** — RichTextBox wrapper for TXT syntax
- **TriggersEditor** — Trigger list editor control

### Operation Opcode ID System

The game's scripting system uses numeric opcodes. `ModOperation.bas` defines them in named constants (e.g., `Call_Script = 1`, `eq = 31`, `troop_add_item = 1529`). These are supplemented by `new.op.ini` which maps additional operation metadata (parameter types, display names, pseudo-code templates).

Operation blocks are stored as `Type_Op_Block` arrays and can be edited in three interchangeable views:
1. **Pseudo-code** (visual tree with dropdowns)
2. **Python** (generated Python-like code)
3. **TXT** (compact internal format: `opcode param_count param1 param2 ...`)

### Mod Detection

The editor auto-discovers mods by reading each mod's `module.ini` for `load_mod_resource` / `load_module_resource` directives. Mod paths are stored in `MnBInfo` (a `Type_Global_Variable_Symbol` instance declared in `ModMain.bas`).

## Important Patterns & Gotchas

### Data Type Definition Pattern

Every game entity type (Troop, Item, Faction, Party, Scene, etc.) follows the same pattern in `ModMain.bas`:
1. A `Type Type_XXX` struct with fields matching the TXT file format
2. A global dynamic array `Public XXX() As Type_XXX`
3. A count variable `Public N_XXX As Long`
4. A `LoadXXXFile(FilePath)` Sub that populates the array
5. A `ReadXXXLine(Text, OutputXXX)` Sub for single-line parsing
6. A `SaveXXXFile(FilePath)` Sub that writes back
7. A `LoadXXXCSVFile(FileName)` Function for localization
8. `SwapXXX` and `ClearXXX` helper Subs
9. A `CurrentXXX` instance and `CurrentXXXID` for tracking active selection

### CSV Loading Performance Issue

All CSV load functions use **nested O(n×m) loops** without early exit — for each CSV line, every mod entry is checked with `LCase$` comparison. For a large mod (2000 items × 2000 CSV lines = 4M string compares). The fix is to build a `Scripting.Dictionary` index first.

### The `GetWord` / `GetWordL` Functions

These are the core low-level parsers. `GetWord` reads from binary file handle `lngHandle` at position `Pointer`, consuming whitespace-delimited tokens byte-by-byte. `GetWordL` does the same from the string variable `txtLine` at position `LinePointer`. Both use `GoTo`-based loops rather than structured `Do While/Loop`. Before calling either, you must set up the global state (`lngHandle`/`Pointer` or `txtLine`/`LinePointer`).

### Global State Coupling

Nearly all functions in `ModMain.bas` depend on global mutable state. You cannot parse a file independently — you must set `lngHandle`, `Pointer`, `MaxPointer` (or `txtLine`, `LinePointer` for text parsing) before calling parsing functions. The `CurrentXXX` variables are set by UI click handlers and read by dozens of functions across different modules.

### 64-bit Integer Handling

Game IDs (items, troops, etc.) are 64-bit numbers internally. `UInt64.bas` provides `Integer64b` — a custom UDT with byte-level access (`by(0)` through `by(7)`). Helper functions: `StrToI64()`, `I64toStrNZ()`, `HexStrToI64()`, `I64toHexStr()`, `I64Add()`, `I64Sub()`, `I64Mul()`, etc. Operations that need 64-bit comparison or arithmetic must use these.

### Language INI Format

Translation files (`.lan.ini`) use Windows INI format. Keys are hardcoded numeric indices. `MdlLanMgr.bas` maintains parallel arrays (`PublicTags(36)`, `PublicEditors(13)`, `PublicTools(5)`, `PublicMsgs(167)`, etc.) — the array size dictates how many keys are read from each INI section. Adding UI text requires updating both the code's array size and all language INI files.

### No Error Recovery in Backup

`ModBackUp.bas` copies files one at a time with `FileCopy`. If any step fails, previously copied files remain in the backup directory with no rollback. The file list is hardcoded in both `SetBackUp` and `RestoreMod` — adding a new mod resource type requires updating both functions.

### Files to Ignore

The source tree contains build artifacts and duplicates that should not be modified:
- `*.OBJ` — compiled VB6 objects (21 files)
- `*.log` — debug log files (`frmCoder.log`, `frmTrigger.log`, `Controls/OpBlockEditor/OpBlockEditor.log`)
- `复件 *.bas`, `复件 *.frm` — "Copy of" backup files
- `ModMain.bas.bak` — 240KB backup of the main module
- `Thumbs.db` — Windows thumbnail cache
- `MnBWarband_Editor.rar` — compressed distribution archive
- `MnBWarband_Editor(Setup).rar` — installer archive
