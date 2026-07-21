# Agent Instructions: Beaver Excel Add-in

VBA Excel Add-in for advanced formatting, cleanup, and reporting.

---

## 🚀 Mandate & Vision
- **Fearless Overhauls**: Empowered and expected to perform high-level architectural changes, module refactoring, and codebase modernization. Unify macros, decouple layers, and optimize performance.
- **High-Level Agency**: Focus on high-level architecture and design system goals; do not let micro-rules block structural improvements.

---

## ⚡ Workflow Commands (`Update.ps1`)
- `pwsh -File .\Update.ps1` — Full build, lint, test execution, & doc update.
- `pwsh -File .\Update.ps1 -Fast` — Fast dev mode (keeps background Excel session alive).
- `pwsh -File .\Update.ps1 -LintOnly -File <path>` — Sub-second syntax check without Excel.
- `pwsh -File .\Update.ps1 -Filter "Test_*"` — Run targeted tests.
- **Diagnostics**: Use temporary `Test_` procedures for bug investigation; clean up once verified.

---

## 🛡️ Non-Negotiable Safety & Quality Rules
1. **Option Explicit**: Required at the start of every module.
2. **Worksheet Qualification**: Always qualify range calls (`ws.Range`, `ws.Cells`) to avoid active sheet bugs.
3. **Null-Safe Property Access**: Check if range properties (e.g. `NumberFormat`, `Font.Name`) are `Null` before type conversion.
4. **Backward Collection Loops**: Iterate backward (`Step -1`) when deleting collection items.
5. **Error Logging**: Wrap entry points and public procedures with `Infra_Error` logging.
6. **Architecture & Prefixes**: Preserve layered architecture (UI, Feature, Infra, Lib, Core) and naming prefixes (`FeatCmd_`, `Infra_`, `UI_`, `Lib_`/`Udf_`, `Test_`).
