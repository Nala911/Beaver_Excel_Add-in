# Agent Instructions: Beaver Excel Add-in

VBA Excel Add-in for advanced formatting, cleanup, and reporting.

---

## 🚀 Architectural Vision & Mandate
- **Fearless Overhauls**: You are fully empowered and expected to perform high-level architectural changes, module refactoring, and codebase modernization. Do not hesitate to restructure code or introduce cleaner patterns (e.g., centralized command dispatch, decoupling layers) when it improves long-term quality or performance.
- **Continuous Improvement**: Proactively look for opportunities to unify fragmented macros, optimize execution speed (using application state toggles and block-array operations), and modularize functionality.

---

## 📂 Workflow
- **Build & Test**: Run `pwsh -File .\Update.ps1` to rebuild, validate syntax/style, run test suites, and auto-update documentation.
- **Controls**: Use parameters on `Update.ps1` (like `-Status`, `-LintOnly`, `-TestCategory`) to speed up cycles.

---

## 📏 Important Rules

1. **Safety & Correctness**:
   - **Option Explicit**: Every module must start with `Option Explicit`.
   - **Explicit Sheet Context**: Always qualify Excel global calls (e.g. `Range`, `Cells`) with a worksheet reference (e.g. `ws.Range`) to prevent unexpected sheet behavior.
   - **Collection Deletion**: Iterate backward (`Step -1`) when deleting items from a collection in a loop.
   - **Type Safety**: Do not perform direct type conversions (e.g., `CStr`, `CLng`) on object/range properties (like `NumberFormat` or `Font.Name`) without checking if the property is `Null` first (e.g., mixed selections return `Null`).

2. **Error & Context Tracking**:
   - Wrap entrypoints and public procedures in structured error handling using the project's error tracker (`Infra_Error`) to ensure application stability and diagnostics.

3. **Style & Structure**:
   - Maintain a clear layered architecture: UI, Feature, Infrastructure, Library, and Core.
   - Use consistent naming prefixes (`FeatCmd_` for features, `Infra_` for infrastructure, `UI_` for user interface, `Udf_`/`Lib_` for library helpers, `Test_` for tests).
