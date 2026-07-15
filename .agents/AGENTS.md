# Agent Instructions: Beaver Excel Add-in

This project is a VBA-based Microsoft Excel Add-in called **Beaver Add-in** that provides advanced formatting, cleanup, highlights, and reporting tools.

---

## 📂 Project Structure & Workflow
- **Rebuilding & Testing**: Run `pwsh -File .\Update.ps1` to lint, compile, and run tests.
- **Source Control**: Only `.bas`, `.cls`, `.frm` source files are tracked. The binary workbook `Beaver Add-in.xlsm` is compiled on demand.
- **Directories**:
  - `Modules/Commands/`: `ICommand` features (prefix/VB_Name: `FeatCmd_`).
  - `Modules/Core/`: Base classes/interfaces (e.g., `AppContainer.cls`, `ICommand.cls`).
  - `Modules/Infrastructure/`: Cross-cutting helpers (prefix/VB_Name: `Infra_`).
  - `Modules/Libraries/`: Helper code (prefix/VB_Name: `Lib_` / `Udf_`).
  - `Modules/Tests/`: Unit tests (prefix/VB_Name: `Test_`).
  - `Modules/UI/`: Forms and Ribbon bindings (prefix/VB_Name: `UI_`).
  - `config.json` & `features.json`: Declarative registries of settings, ribbon controls, and features.

---

## 📏 Development Rules & Linter Constraints
Every code file is automatically verified by `.build/Linter.ps1`. Adhere to the following strictly:

### 1. Mandatory File Headers
- Every file must start with `Option Explicit` followed by:
  ```vba
  ' @Module: ModuleName
  ' @Category: Core | Infrastructure | Feature | Library | UI
  ' @Description: Description of the module.
  ```

### 2. Context Tracking & Error Handling
- Every public procedure (except events/infra/libs) and command execution must track context and handle errors:
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
- **Range Formulas**: Always use `Range.Formula2` instead of `Range.Formula` to prevent dynamic array spill errors.
- **Direct Conversions**: Do not call conversion functions (e.g., `CStr`, `CLng`) directly on range properties (e.g., `NumberFormat`, `Font.Name`, `Font.Size`) without first checking for `IsNull`. Mixed ranges return `Null`, causing Error 94.
- **Loop Deletion**: When modifying or deleting items in a collection inside a loop, iterate backwards (`For i = count To 1 Step -1`).
- **Explicit References**: Always qualify references to Excel globals (`Range`, `Cells`, `Rows`, `Columns`) with a sheet variable (e.g., `ws.Range`) to prevent active sheet bugs.
- **Localization**: Do not hardcode localized sheet names (e.g., `"Sheet1"`). Use constants or dynamic discovery.

### 4. Standardized Naming Conventions
- **Variable Casing**: camelCase for local variables and internal parameters.
- **Public APIs/Properties**: PascalCase or Snake_Case.
- **Files/VB_Name prefixes**: `FeatCmd_` (Features), `Infra_` (Infrastructure), `UI_` (UI), `Udf_` (UDFs), `Test_` (Tests).
