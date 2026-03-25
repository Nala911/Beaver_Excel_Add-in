# Beaver Add-in - AGENTS Guide

Read this before changing VBA, Ribbon XML, or project config. This file is the project skill context for agents working in this repo.

Last updated: 2026-03-24
Current config version: 2.0.64

## Purpose

Beaver is an Excel VBA add-in for fast workbook operations:

- range formatting and paste-format helpers
- data cleanup and date normalization
- formula manipulation and formula-to-value conversion
- link breaking, workbook duplication, and range export
- workbook hotkeys and Ribbon-driven entry points
- worksheet UDFs: `XFilter` and `XUnpivot`
- lightweight unit testing framework

## Source Of Truth

Treat the exported text files in this repo as authoritative:

- `Modules\*.bas` and `Modules\*.cls` are the editable VBA source.
- `ThisWorkbook.cls` is the authoritative workbook event module on disk.
- `features.json` is the authoritative feature manifest for Ribbon controls, release tiers, icons, and hotkeys.
- `ribbon.xml` is generated from `features.json` by `Update.ps1`.
- `config.json` is the authoritative configuration for add-in identity and UI constants; `Update.ps1` synchronizes its `Hotkeys`, `Icons`, and `FeatureFlags` sections from `features.json`.
- `Beaver Add-in.xlsm` is the compiled workbook artifact that `Update.ps1` syncs into.

Do not make manual edits inside the VBA editor and leave disk files unchanged. Edit the exported files first, then run `.\Update.ps1`.

## Do Not Touch

- `Beaver Add-in.xlsm` directly ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â always use `Update.ps1`
- `Lib_JsonConverter.bas` ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â third-party library, do not modify
- `Update.ps1` unless the build pipeline itself is the task
- `.frx` binary form files ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â these are auto-generated

## Project Layout

```text
Excel Add-in\
|- Beaver Add-in.xlsm
|- ThisWorkbook.cls
|- ribbon.xml
|- features.json
|- config.json
|- Update.ps1
|- AI_GUIDE.md
\- Modules\
   |- Feat_*.bas        feature macros
   |- Infra_*.bas/.cls  shared infrastructure and typed request/context classes
   |- UI_Ribbon.bas     Ribbon callbacks
   |- Lib_*.bas         libraries and UDFs
   \- Lib_Tests.bas     unit testing framework
```

## Runtime Architecture

There are three main entry paths:

1. Workbook lifecycle:
   `ThisWorkbook.Workbook_Open` registers hotkeys via `Infra_Hotkeys.RegisterHotkeys`.
   `ThisWorkbook.Workbook_BeforeClose` unregisters them.

2. Ribbon callbacks:
   `UI_Ribbon.bas` contains the `onAction` procedures referenced by `ribbon.xml`.
   Those callbacks delegate into `Feat_*` modules.

3. Hotkeys:
   `Infra_Hotkeys` reads definitions from `config.json` and binds them with `Application.OnKey`.

Shared support layers:

- `Infra_AppStateGuard` captures and restores `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, and `Calculation`.
- `Infra_AppState` provides selection guards, context capture, and Desktop-path resolution.
- `Infra_Config` loads `config.json` into a typed `Infra_ConfigModel`.
- `Infra_Error` centralizes breadcrumb tracking, environment snapshotting, user-visible error dialogs (with automatic clipboard copying), and Excel failsafe reset.
- `Infra_ContextTracker` (RAII) handles automated `PushContext`/`PopContext` via `Infra_Error.Track`.
- `Infra_Undo` provides custom undo management for macro-driven changes using a hidden `_BeaverUndo` sheet.
- `Infra_Progress` provides centralized progress reporting via the Excel status bar.
- Typed request/context classes exist for multi-step actions:
  `Infra_ActionContext`, `Infra_CleanDataRequest`, `Infra_ExportRequest`, `Infra_ErrorContext`.

## Current Feature Surface

Ribbon-backed features:

- `Feat_MergeFormulas.MergeFormulas`
  Inlines a precedent into a dependent formula using `.Formula2`, including spill-reference handling.
- `Feat_WrapSelectedRange.WrapSelectionWithFormula`
  Wraps selected formulas or constants using a user template and `[value]` placeholder.
- `Feat_MakeItStatic.StaticSheetWorkbook`
  Converts formulas to values across active sheet or whole workbook, with spill-aware handling.
- `Feat_CleanData.CleanData`
  Cleans text constants only, using `Trim` and `Clean`, across selection, sheet, or workbook.
- `Feat_BreakExternalLinks.BreakExternalLinks`
  Detects and removes workbook links, connections, external named ranges, external formulas, pivot output, and non-range tables.
- `Feat_DateConversion.ConvertTextToProperDate`
  Converts a single selected column of text dates using a user-provided target month.
- `Feat_Duplicate.Duplicate`
  Copies all sheets to a new macro-free `.xlsx` on Desktop.
- `Feat_ExportImageOrPdf.Export`
  Exports the selected range or `UsedRange` as Desktop PNG or PDF.
- `Feat_ToggleFullScreen.ToggleFullScreen`
  Toggles worksheet UI chrome such as headings, tabs, scrollbars, and formula bar.
- `Infra_Hotkeys.ShowHotkeysHelp`
  Builds a help sheet in the active workbook.

Hotkey-only or config-driven features:

- `Feat_FormatRange.FormatSelectedRange`
  Applies bulk table formatting, unmerges cells, autofits, and applies number/date formats.
- `Feat_FormatRange.ApplyCustomNumberFormat`
- `Feat_FormatRange.PasteFormat`
- `Feat_CreateSheet.CreateNamedSheet`
- `Feat_FillDown.FillDown`
- `Feat_FilterByCell.FilterBySelectedCell`
- `Feat_BackspaceDelete.Backspace`
- `Feat_BackspaceDelete.Delete`
- `Feat_MakeItStatic.MakePermanent`

Worksheet UDFs:

- `Lib_XFilterFunction.XFilter`
  Performs intersection or difference using dictionary-backed lookups and returns a spill array.
  Arguments: `XFilter(Range_A, Range_B, code_number)` ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â code 1 = intersection, 2 = difference.

- `Lib_XUnpivotFunction.XUnpivot`
  Transforms a wide (crosstab) range into a long (normalised) format and returns a spill array.
  Arguments: `XUnpivot(SourceRange, FixedColumnsCount, [IgnoreBlanks], [AttributeHeader], [ValueHeader])`
  Returns `#NULL!` when the result set is empty, `#VALUE!` on input errors.
  The first row of `SourceRange` is treated as a header row and is preserved in the output.

Testing:

- `Lib_Tests.RunAllTests`
  Executes all registered unit test suites and reports results to the Immediate Window.

## Current Ribbon And Hotkey Contract

Ribbon groups defined in `ribbon.xml`:

- `Formatting`: merge formulas, wrap formula, static sheet/workbook
- `Data Tools`: clean data, break links, fix dates
- `File Actions`: duplicate, export range
- `Workspace`: focus mode
- `Support`: hotkeys help

Hotkeys currently defined in `config.json`:

- `^+4` -> `Feat_FormatRange.ApplyCustomNumberFormat`
- `^+p` -> `Feat_MakeItStatic.MakePermanent`
- `+{F11}` -> `Feat_CreateSheet.CreateNamedSheet`
- `^%{DOWN}` -> `Feat_FillDown.FillDown`
- `^+f` -> `Feat_FilterByCell.FilterBySelectedCell`
- `^+m` -> `Feat_FormatRange.PasteFormat`
- `^+q` -> `Feat_FormatRange.FormatSelectedRange`
- `{BACKSPACE}` -> `Feat_BackspaceDelete.Backspace`
- `{DELETE}` -> `Feat_BackspaceDelete.Delete`

If you add a new shortcut, put it in `config.json`. Do not hardcode new `Application.OnKey` bindings in feature modules.

## Config Contract

`config.json` currently contains:

- `Hotkeys`
- `UIConstants`
- `AddinIdentity`
- `Icons`

### config.json Schema Snapshot

```json
{
  "AddinIdentity": { "Name": "...", "Version": "..." },
  "UIConstants":   { 
      "DefaultFontName": "...",
      "DefaultFontSize": 10,
      "HeaderFontSize": 11,
      "DefaultNumberFormat": "...",
      "DefaultDateFormat": "...", 
      "DisplayDateFormat": "...",
      "ColumnWidthThreshold": 40,
      "MaxColumnWidth": 25,
      "HeaderColor": "#AEAAAA",
      "DefaultExportScale": 3,
      "MaxExportScale": 10
  },
  "Icons":         { "BtnCleanData": "ImageMsoName" },
  "Hotkeys":       [ { "Key": "^+p", "Macro": "...", "Description": "..." } ]
}
```

`Infra_Config` exposes typed getters for frequently used values, including:

- add-in name/version
- default and display date formats
- default number format
- header color and font sizes
- default font name/size
- export scale limits
- column width thresholds
- Ribbon icon lookup

If a new shared constant is needed, add it to `config.json`, then expose it via `Infra_Config` and `Infra_ConfigModel`.

## Error Handling Pattern

Public entry points should use the centralized error stack and RAII context tracker:

```vba
Public Sub MyMacro()
    Dim tracker As Object: Set tracker = Infra_Error.Track("MyMacro")
    Dim guard As New Infra_AppStateGuard
    On Error GoTo ErrHandler

    ' work
    ' (Optional: use Infra_Undo.SaveState Target, "Action Name")
    ' (Optional: use Infra_Progress.StartProgress "Title", TotalSteps)

CleanExit:
    ' Progress and Guard clean up automatically via RAII/Destructors
    Exit Sub

ErrHandler:
    Infra_Error.HandleError "MyMacro", Err
End Sub
```

Rules:

- Use `Infra_Error.Track` at the start of every public macro, Ribbon callback, and workbook event.
- Use `Infra_AppStateGuard` for macros that alter Excel application state (ScreenUpdating, etc.).
- Use `Infra_Undo.SaveState` BEFORE modifying a range to enable custom Undo support.
- Use `Infra_Progress` for any operation that may take more than a second to provide UI feedback.
- Prefer a single `CleanExit` label when there are multiple early exits.
- Call `Infra_Error.HandleError` in the handler; do not build ad hoc error dialogs unless the code has a very specific UX reason.

## VBA Coding Rules

- `Option Explicit` is mandatory in every `.bas` and `.cls` file.
- Every VBA source file must include the metadata header:

```vba
' @Module: ModuleName
' @Category: Infrastructure/Feature/UI/Library
' @Description: Short description
' @ManagedBy: BeaverAddin Agent
' @Dependencies: Dependency list
```

- Keep public entry modules thin. Put reusable logic in helpers or typed request/context classes when a feature has multiple decision stages.
- Prefer `CaptureActionContext()` when a feature depends on active workbook, sheet, selection, or active cell.
- Prefer array-based range processing over cell-by-cell loops when mutating many cells.
- Use `.Formula2` when working with modern formulas or spill-aware logic.
- Preserve user-facing configuration in `config.json`, not in scattered literals.
- **NEVER** modify `_BeaverUndo` or `_BeaverTests` (if they exist) manually.

## Build And Sync Workflow

Standard workflow:

1. Edit files on disk.
2. Run `.\Update.ps1`.
3. Fix any validation, compile, Ribbon, or runtime-test failures.

What `Update.ps1` currently does:

- validates `ribbon.xml`
- synchronizes `features.json` into `ribbon.xml` and config hotkeys/icons/feature flags
- generates `Modules\Lib_TestManifest.bas` from `Public Sub Test_*` procedures
- verifies Ribbon callbacks exist in exported VBA
- performs regex-based VBA structure checks
- enforces `Option Explicit` and metadata headers
- closes the workbook if Excel has it locked
- removes managed VBA components from the workbook
- imports all `.bas` and `.cls` files from `Modules\`
- replaces `ThisWorkbook` code from `ThisWorkbook.cls`
- compiles the VBA project through Excel/VBE automation
- optionally increments the patch/build version in `config.json` when `-BumpVersion` is passed
- saves the workbook
- injects `ribbon.xml` into the `.xlsm` archive as `customUI/customUI14.xml`
- runs runtime smoke tests against each Ribbon callback unless `-SkipRuntimeTests` is passed

### Runtime Smoke Tests (Update.ps1)

`Update.ps1` opens the workbook, validates the generated Ribbon, and runs `Lib_Tests.RunAllTests`.
Use `-SkipRuntimeTests` only when deliberately deferring runtime validation.

Prerequisites:

- Excel must allow "Trust access to the VBA project object model".
- The environment must support Excel COM automation.

## Agent Guidance

- Read the exported module you are changing and its immediate infra dependencies first.
- Keep `ribbon.xml`, `config.json`, and VBA callbacks aligned. A Ribbon button without a callback, or a callback without a module entry point, will fail validation.
- Prefer editing `features.json` over hand-editing generated Ribbon or synced config sections.
- If you add a new Ribbon control, also add its icon mapping to `config.json` if it uses `Ribbon_GetIcon`.
- If you add a new config field, update both `Infra_ConfigModel` and `Infra_Config`.
- Do not assume the workbook internals are authoritative over disk exports. The repo files are the working source.
- After meaningful changes, run `.\Update.ps1`. Use `-SkipRuntimeTests` only when runtime smoke tests are intentionally being deferred.

## Agent Checklist (run mentally before every edit)

- [ ] Read the target module + its `@Dependencies` first
- [ ] If touching `ribbon.xml` ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ verify `onAction` name exists in `UI_Ribbon.bas`
- [ ] If adding a Ribbon button ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ add its icon to `config.json` `Icons` section
- [ ] If adding a public `Sub` ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ follow the `Infra_Error.Track` error pattern
- [ ] If adding a config value ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ update `Infra_ConfigModel` AND `Infra_Config`
- [ ] If mutating a range ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ use `Infra_Undo.SaveState`
- [ ] If a task takes time ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ use `Infra_Progress`
- [ ] Run `.\Update.ps1` after changes and fix all reported failures

## Known Gotchas

- **`.Formula2` vs `.Formula`**: Always use `.Formula2` for modern/spill formulas. Using `.Formula` will inject `@` implicit intersection operators.
- **`.frx` files**: Never edit manually. They are binary; only the VBE regenerates them.
- **`Infra_AppStateGuard`**: Declare as `New` inline (`Dim guard As New ...`), not as a separate `Set`. The class destructor runs on `Exit Sub` to restore app state.
- **`Infra_Error.Track`**: Assign to an object variable (`Dim tracker As Object: Set tracker = ...`) to ensure the RAII destructor runs at the correct time (when the variable goes out of scope at `Exit Sub`).
- **`config.json` version**: `Update.ps1` auto-increments the patch version. Do not manually set it or your increment will be overwritten.
- **Ribbon changes**: Ribbon edits require Excel restart (or workbook close/reopen) to reflect XML changes.
- **Manifest-driven UI**: `BtnDummyHello` is marked `dev` in `features.json`; pass `-IncludeDevFeatures` to `Update.ps1` if you intentionally want developer-only controls in the generated Ribbon/config.
