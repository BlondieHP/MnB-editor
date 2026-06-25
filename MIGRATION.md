# Platform Migration Strategy

## Why Migrate?

Visual Basic 6.0 reached end of support in **2008**. The IDE does not run reliably on Windows 10/11. The runtime (msvbvm60.dll) is still shipped with Windows but receives no updates. Key risks:

| Risk | Impact |
|------|--------|
| VB6 IDE won't install on modern Windows | Cannot build or modify the project |
| No 64-bit support | Cannot handle large mod files (>2GB) |
| No Unicode support in native controls | Chinese text rendering issues |
| No modern security features | ASLR, DEP, code signing all unavailable |
| No automated testing framework | Every change requires manual testing |
| No package manager / dependency management | All dependencies manually tracked |

## Target Platform Comparison

### Option A: .NET 8 + Windows Forms (C#)

| Criterion | Rating | Notes |
|-----------|--------|-------|
| Migration effort | Medium | VB6→.NET has tooling support; similar event-driven paradigm |
| Performance | Excellent | JIT-compiled, 64-bit, proper generics |
| UI fidelity | Good | WinForms is the direct successor to VB6 forms |
| Cross-platform | Partial | Windows primary; Linux/Mac via Mono (limited) |
| MOD community | Good | .NET is popular for game tools |
| Long-term support | Excellent | .NET 8 LTS supported until 2026; .NET 10+ planned |

### Option B: Python + PyQt6

| Criterion | Rating | Notes |
|-----------|--------|-------|
| Migration effort | High | Complete rewrite; VB6 and Python are very different |
| Performance | Moderate | Python is slower but file I/O dominates in this case |
| UI fidelity | Good | PyQt provides rich controls equivalent to VB6 |
| Cross-platform | Excellent | Windows, Linux, macOS |
| MOD community | Excellent | Most M&B modders use Python (the game's scripting is Python-like) |
| Long-term support | Good | Python is community-maintained; PyQt has commercial backing |

### Option C: Web (Electron / Tauri + React/Vue)

| Criterion | Rating | Notes |
|-----------|--------|-------|
| Migration effort | Very High | Complete rewrite with different paradigm |
| Performance | Moderate | Web stack overhead but adequate for data editing |
| UI fidelity | Flexible | Can exceed VB6 UI quality |
| Cross-platform | Excellent | Windows, Linux, macOS, web |
| MOD community | Moderate | Web tools gaining traction |
| Long-term support | Good | Depends on framework choice |

## Recommended Path: .NET 8 + Windows Forms (C#)

**.NET is the most natural migration target** because:
1. VB6 and WinForms share the same event-driven, control-based paradigm
2. Microsoft provides a VB6→.NET upgrade wizard (though manual rewrite is better)
3. The DataRepository pattern already established maps directly to .NET properties/events
4. C# can call native Win32 APIs for any remaining Windows-specific functionality
5. .NET's `Span<T>` and `Memory<T>` provide zero-allocation binary parsing — perfect for the game's TXT file format

## Migration Architecture

```
┌──────────────────────────────────────────────┐
│                  UI Layer                     │
│  (WinForms / WPF / PyQt / React)             │
│  frmItems, frmTroops, frmMain, ...           │
│  Binds to DataRepository, handles events     │
├──────────────────────────────────────────────┤
│              DataRepository                   │
│  In-memory data store with change events     │
│  Troop[], Item[], Faction[], ...             │
│  ObservableCollection<T> or custom events    │
├──────────────────────────────────────────────┤
│           File I/O Layer                      │
│  ┌──────────┐  ┌──────────┐  ┌───────────┐  │
│  │ TXT      │  │ CSV      │  │ Python    │  │
│  │ Parser   │  │ Loader   │  │ Exporter  │  │
│  └──────────┘  └──────────┘  └───────────┘  │
│  (GetWord,     (Dict-based,  (Operation    │  │
│   PutWord,      O(n) lookup)  code export) │  │
│   buffer I/O)                              │  │
├──────────────────────────────────────────────┤
│              Shared Models                    │
│  Troop, Item, Faction, Party, Scene, ...     │
│  (C# records / Python dataclasses)           │
└──────────────────────────────────────────────┘
```

## Step-by-Step Migration Plan

### Phase 1: Model Extraction (1-2 weeks)
- Define all data types as C# `record` or `record struct` types
- Port Constants.bas → `Constants.cs` (static class with const values)
- Port Models.bas → `Models.cs` (record definitions)
- Write round-trip serialization tests: load TXT → serialize → load → compare

### Phase 2: File I/O Port (1-2 weeks)
- Port FileParser.bas → `FileParser.cs` using `Span<byte>` for zero-allocation parsing
- Port CsvLoader.bas → `CsvLoader.cs` using `Dictionary<string, int>` (already O(n))
- Port FileLoader.bas → `FileLoader.cs` — one loader per entity type
- Port FileSaver.bas → `FileSaver.cs` — one saver per entity type
- **Write unit tests for every Load/Save function** (this is the critical validation layer)

### Phase 3: DataRepository Port (1 week)
- Port DataRepository.cls → `DataRepository.cs`
- Use `ObservableCollection<T>` for UI binding
- Implement `INotifyPropertyChanged` for all properties
- Write integration tests for all CRUD operations

### Phase 4: UI Port (3-4 weeks)
- Port forms one at a time, starting with the simplest
- Order: frmAbout → frmColor → frmBox → frmFactions → frmItems → frmTroops → frmParties → frmMain
- Each form inherits from a base `EditorForm` class with common functionality
- Use data binding: `textBox.DataBindings.Add("Text", repo, "CurrentItem.Name")`
- The `ActiveString` function becomes a C# interpolated string or `string.Format`

### Phase 5: Integration & Polish (1-2 weeks)
- Port Init.bas → `AppInitializer.cs` (dependency injection setup)
- Port Utils.bas → `Utilities.cs` (extension methods on model types)
- Port ModPython.bas → `PythonExporter.cs`
- End-to-end testing with real mod files
- Performance comparison against VB6 original

### Phase 6: Deployment (1 week)
- Package as single-file executable: `dotnet publish -p:PublishSingleFile=true`
- Code signing certificate
- Auto-update mechanism (ClickOnce or custom)

## Key Migration Risks & Mitigations

| Risk | Likelihood | Mitigation |
|------|-----------|------------|
| File format parsing differences | Medium | Write round-trip tests: load original TXT → save → binary compare with original |
| GBK encoding issues | Medium | .NET supports all code pages via `Encoding.GetEncoding(936)`; test with Chinese mod files |
| UI layout differences | High | Accept minor visual differences; focus on functional equivalence |
| Missing VB6 built-in functions | Low | Map `Val()`, `Mid()`, `Chr()` to .NET equivalents; write adapter if needed |
| 64-bit integer handling | Low | .NET has `Int64`/`UInt64` natively; remove Integer64b entirely |

## Python Alternative Path

If Python is preferred, the same layered architecture applies but with different tools:

```
Models:     @dataclass classes
Parser:     memoryview + struct module for binary parsing
Storage:    DataRepository with Observable pattern (custom or RxPy)
UI:         PyQt6 with QDataWidgetMapper for form binding
Packaging:  PyInstaller for single .exe delivery
```

Python advantages for this specific project:
- M&B modding community already uses Python
- The game's module system uses Python-like syntax
- Easier to add scripting/plugin support
- Cross-platform (some modders use Linux/Wine)

## Immediate Next Steps (can be done in current VB6 codebase)

1. **Adopt DataRepository.cls** in all forms — replaces direct global array access
2. **Standardize error handling** — all forms use the same pattern
3. **Extract shared UI logic** — form base class with common methods
4. **Write data validation rules** — centralize in DataRepository property setters
5. **Create a test data set** — a known-good mod that exercises all features
