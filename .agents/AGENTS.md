# Agent Instructions: Beaver Excel Add-in

VBA Excel Add-in for advanced formatting, cleanup, and reporting.

---

## 🚀 Architecture & Mandate
- **System Architecture**: Consult [`ARCHITECTURE.md`](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Excel%20Add-in/ARCHITECTURE.md) for subsystem maps, sequence diagrams, layer dependency rules (`UI` → `Feature` → `Infra` → `Lib` → `Core`), and entry point mappings before refactoring or adding features.
- **Fearless Overhauls**: Empowered to perform high-level architectural changes, module refactoring, and codebase modernization without being blocked by micro-rules.

---

## ⚡ Workflow Commands (`Update.ps1`)
- `pwsh -File .\Update.ps1` — Standard build, lint, & test execution (skips `.xlam` export for speed).
- `pwsh -File .\Update.ps1 -ExportAddin` — Full build, test execution, & export compiled `Beaver.xlam` add-in binary.
- `pwsh -File .\Update.ps1 -Fast` — Fast dev mode (keeps background Excel session alive).
- `pwsh -File .\Update.ps1 -LintOnly -File <path>` — Sub-second syntax check without Excel.
- `pwsh -File .\Update.ps1 -Filter "Test_*"` — Run targeted tests.
- **Diagnostics**: Use temporary `Test_` procedures for bug investigation; clean up once verified.

---

## 🛡️ Non-Negotiable Rules
1. **Option Explicit**: Line 1 of every `.bas`, `.cls`, `.frm` module.
2. **Worksheet Qualification**: Always qualify range calls (`ws.Range`, `ws.Cells`).
3. **Null-Safe Property Access**: Check range properties for `Null` before type casting.
4. **Backward Collection Loops**: Iterate backward (`Step -1`) when deleting collection items.
5. **Error Logging**: Wrap entry points and public procedures with `Infra_Error` logging.
6. **Architecture & Prefixes**: Preserve layered architecture and naming prefixes (`FeatCmd_`, `Infra_`, `UI_`, `Lib_`/`Udf_`, `Test_`).
