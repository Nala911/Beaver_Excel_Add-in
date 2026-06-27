# Beaver Add-in - Workspace Rules for AI Agents

These rules apply specifically to all AI agents executing tasks in this workspace. They complement the global rules and establish strict limits on how documentation is updated.

---

## Guidelines for Modifying GEMINI.md

To prevent context window bloat and maintain agent performance, the file [GEMINI.md](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Excel%20Add-in/GEMINI.md) must remain a high-level operational guardrail playbook. Do not allow it to expand into a verbose reference manual.

### 1. Document Size Constraint
* **Strict Limit**: [GEMINI.md](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Excel%20Add-in/GEMINI.md) must not exceed **120 lines**. If it grows past this limit, you must prune or consolidate sections.

### 2. What to NEVER Add to GEMINI.md
* **No command lists or control catalogs**: Do not list command names (`FeatCmd_*`), ribbon control IDs (`Btn*`), or hotkey macros. Pointers to `features.json` and directory paths are sufficient.
* **No configuration snapshots**: Do not copy parts of `config.json` or its schema. Direct agents to read the file.
* **No long code boilerplates**: Do not include large VBA code blocks or error-handling examples. Pointers to reference modules (e.g. `FeatCmd_HelloWorld`) are sufficient.
* **No feature-specific internal explanations**: Implementation details and internal mechanics of specific features must be documented via comments inside the code modules, not in the playbook.

### 3. What is ALLOWED to be updated in GEMINI.md
* Changes to the unified build/sync commands (e.g. `Update.ps1` params).
* Major changes to the core system architecture (e.g. if the core execution pipeline transitions from command-based to event-driven).
* Core UI design patterns (e.g. adding new shared UserForm rules).

### 4. Review Process
* Any modification to [GEMINI.md](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Excel%20Add-in/GEMINI.md) must be proposed in the **Implementation Plan** and explicitly approved by the user.
