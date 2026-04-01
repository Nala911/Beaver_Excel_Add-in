# Beaver Add-in - Gemini Context

This project is a high-performance Excel VBA add-in designed for rapid workbook operations, data cleanup, and formula manipulation. It follows a "disk-first" development workflow where authoritative source files are maintained as text files and synchronized into a compiled `.xlsm` artifact.

## Project Overview

- **Name:** Beaver Add-in
- **Technologies:** Excel VBA, PowerShell (Build/Sync), XML (Ribbon UI), JSON (Configuration/Manifests).
- **Architecture:** 
  - **Source of Truth:** Authoritative code resides in `Modules/` (`.bas`, `.cls`) and root-level config files (`features.json`, `config.json`).
  - **Ribbon UI:** Defined in `features.json` and generated into `ribbon.xml`.
  - **Build System:** `Update.ps1` automates the synchronization of source files into the `Beaver Add-in.xlsm` workbook.
  - **Core Infrastructure:** Centralized error handling (`Infra_Error`), application state management (`Infra_AppStateGuard`), and custom undo support (`Infra_Undo`).

## Building and Running

### Build and Sync
The project uses PowerShell to sync disk changes into the Excel workbook. Always run this after modifying source files.

```powershell
# Standard sync and validation
.\Update.ps1

# Sync with version increment
.\Update.ps1 -BumpVersion

# Sync including developer-only features
.\Update.ps1 -IncludeDevFeatures

# Skip runtime smoke tests (faster)
.\Update.ps1 -SkipRuntimeTests
```

### Running the Add-in
1. Open `Beaver Add-in.xlsm` in Excel.
2. The "BEAVER" tab will appear in the Ribbon.
3. Use Hotkeys (e.g., `Ctrl+Shift+Q` for formatting) or Ribbon buttons to trigger features.

### Testing
- **Automated:** `Update.ps1` runs smoke tests by default.
- **Manual:** Run `Lib_Tests.RunAllTests` from the VBA Immediate Window.

## Development Conventions

### Authoritative Files
**DO NOT** edit code directly in the VBA Editor (VBE) without exporting it back to disk. The disk files are the source of truth.
- `Modules/*.bas` and `Modules/*.cls`: VBA source.
- `features.json`: Source for Ribbon controls, icons, and hotkey assignments.
- `config.json`: Identity, UI constants, and synced feature flags.

### VBA Coding Standards
- `Option Explicit` is mandatory in all modules.
- **Metadata Header:** Every file must include:
  ```vba
  ' @Module: ModuleName
  ' @Category: Infrastructure/Feature/UI/Library
  ' @Description: Short description
  ' @ManagedBy: BeaverAddin Agent
  ' @Dependencies: Dependency list
  ```
- **Error Handling:** Use the RAII pattern for entry points:
  ```vba
  Public Sub MyMacro()
      Dim tracker As Object: Set tracker = Infra_Error.Track("MyMacro")
      Dim guard As New Infra_AppStateGuard
      On Error GoTo ErrHandler
      ' ... logic ...
  CleanExit:
      Exit Sub
  ErrHandler:
      Infra_Error.HandleError "MyMacro", Err
  End Sub
  ```
- **Undo Support:** Call `Infra_Undo.SaveState` before range mutations.
- **Formulas:** Use `.Formula2` for modern Excel spill-range compatibility.

## Key Files

- `Update.ps1`: The build and synchronization script.
- `features.json`: The master manifest for the UI and hotkeys.
- `config.json`: Centralized configuration and constants.
- `AGENTS.md`: Detailed technical guide for AI agents and developers.
- `Modules/Infra_*.bas`: Core framework and infrastructure.
- `Modules/Feat_*.bas`: Feature-specific implementations.
- `Modules/UI_Ribbon.bas`: Ribbon callback handlers.
