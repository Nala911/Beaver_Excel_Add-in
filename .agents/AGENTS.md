# Agent Instructions: Beaver Excel Add-in

VBA Excel Add-in for advanced formatting, cleanup, and reporting.

---

## 📂 Project Structure & Workflow
- **Rebuilding & Testing**: Run `pwsh -File .\Update.ps1`.
- **Architecture Map**: Automatically regenerated on successful build/test execution by `Update.ps1` when changes are detected. Can also be manually forced with `pwsh -File .\Update.ps1 -GenerateDocs` or `pwsh -File .build\GenerateArchitectureMap.ps1`.
- **Architecture Reference**: Check [ARCHITECTURE.md](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Excel Add-in/ARCHITECTURE.md) for module layering, dependencies, and mapping enums before any refactoring.
- **Source Control**: Only `.bas`, `.cls`, `.frm`, and the auto-generated `ARCHITECTURE.md` are tracked.

### 🛠️ Developer & Agent Command-Line Controls
Run `Update.ps1` at the root with these agent-centric switches for faster cycles:
- `pwsh -File .\Update.ps1 -Status` — Get repository diagnostics, module count, and lint health report.
- `pwsh -File .\Update.ps1 -LintOnly [-AutoFix]` — Run quick syntax and style validation (bypassing Excel entirely).
- `pwsh -File .\Update.ps1 -TestCategory UI|Feature|Infrastructure|Core` — Test only a specific layer of code.
- `pwsh -File .\Update.ps1 -GenerateDocs` — Force regeneration of `ARCHITECTURE.md` Mermaid map.
- `pwsh -File .\Update.ps1 -SkipDocs` — Skip documentation auto-generation.

---

## 📏 Development & Linter Rules (Verified by `.build/Linter.ps1`)

### 1. Mandatory Headers
- Must start with `Option Explicit` followed by:
  ```vba
  ' @Module: ModuleName
  ' @Category: Core | Infrastructure | Feature | Library | UI
  ' @Description: Description of the module.
  ```

### 2. Context Tracking & Error Handling
- Public procedures (except events/infra/libs) and commands must track context:
  ```vba
  Public Sub ProcedureName()
      Dim tracker As Object: Set tracker = Infra_Error.Track("ProcedureName")
      On Error GoTo ErrHandler
      ' Implementation here
  CleanExit:
      Exit Sub
  ErrHandler:
      Infra_Error.HandleError "ProcedureName", Err
      Resume CleanExit
  End Sub
  ```

### 3. Safety & Design Rules
- **Range Formulas**: Use `Range.Formula2` instead of `Range.Formula`.
- **Direct Conversions**: Do not call `CStr`, `CLng`, etc. directly on Range properties (e.g., `NumberFormat`, `Font.Name`) without checking `IsNull` first.
- **Loop Deletion**: Iterate backwards when deleting from collections (`For i = count To 1 Step -1`).
- **Explicit References**: Qualify all Excel globals (`Range`, `Cells`, etc.) with a sheet variable (e.g., `ws.Range`).
- **Conventions**: camelCase for local variables; PascalCase/Snake_Case for public APIs; prefixes: `FeatCmd_` (Features), `Infra_` (Infrastructure), `UI_` (UI), `Udf_`/`Lib_` (Libraries), `Test_` (Tests).
