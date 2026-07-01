# Agent Instructions: Beaver Excel Add-in

This project is a VBA-based Microsoft Excel Add-in called **Beaver Add-in** that provides advanced formatting, cleanup, highlights, and reporting tools.

---

## 📂 Project Structure

- `Modules/Commands/`: Implementations of `ICommand` (e.g., `FeatCmd_*.cls`).
- `Modules/Core/`: Base classes/interfaces (e.g., `AppContainer.cls`, `ICommand.cls`, `IConfig.cls`).
- `Modules/Infrastructure/`: Cross-cutting concerns (e.g., `Infra_Error.cls`, `Infra_Undo.bas`, `Infra_Config.cls`).
- `Modules/Libraries/`: Helper code (e.g., `Lib_JsonConverter.bas`, custom UDFs).
- `Modules/Tests/`: Unit tests (e.g., `Lib_Tests_Feat_*.bas`, `Lib_Tests.bas`).
- `Modules/UI/`: Forms and Ribbon bindings (e.g., `UI_Ribbon.bas`).
- `config.json`: Constant and Hotkey declarations.
- `features.json`: Declarative registry of ribbon controls, hotkeys, and features.
- `Update.ps1`: Main developer script. Run `pwsh -File .\Update.ps1` to lint, compile, and run tests.

---

## 🛠️ Workflow & Build Pipeline

1. **Source Control**: Only `.bas`, `.cls`, `.frm` source files are tracked. The binary workbook `Beaver Add-in.xlsm` is compiled on demand.
2. **Rebuilding**: After making edits to any source files, run:
   ```powershell
   pwsh -File .\Update.ps1
   ```
   This will import all changed files, compile the VBA project, and run unit tests.

---

## 📏 Development Rules & Linter Constraints

Every code file is automatically verified by `.build/Linter.ps1`. Adhere to the following strictly:

### 1. Mandatory File Headers
- Every file must start with `Option Explicit`.
- Every file must have a metadata header comment matching:
  ```vba
  ' @Module: ModuleName
  ' @Category: Infrastructure | Feature | Library | UI
  ' @Description: Description of the module.
  ```

### 2. Context Tracking & Error Handling
- Every public procedure and command execution must track context and handle errors using the project's standard patterns:
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

### 3. Spill-Safe Range Formulas
- Always use `Range.Formula2` instead of `Range.Formula` to prevent dynamic array spill errors.

### 4. Direct Conversion Safe Guards
- Do not call conversion functions (e.g., `CStr`, `CLng`) directly on range properties (`NumberFormat`, `Font.Name`, `Font.Size`) without first checking for `IsNull`. Mixed ranges return `Null`, causing Run-time Error 94.

### 5. Loop Deletion / Collection Mutation
- When modifying or deleting items in a collection/array inside a loop, always iterate backwards:
  ```vba
  For i = count To 1 Step -1
  ```

### 6. Explicit References
- Always qualify references to Excel globals like `Range`, `Cells`, `Rows`, and `Columns` with a sheet variable (e.g., `ws.Range`) to prevent `ActiveSheet` selection bugs.

### 7. Sheet Localization
- Do not hardcode localized sheet names (e.g., `"Sheet1"`), which fail in non-English Excel environments. Use constants or dynamic discovery.
