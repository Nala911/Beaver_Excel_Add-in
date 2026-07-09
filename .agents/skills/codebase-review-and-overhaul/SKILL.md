---
name: Codebase Review and Overhaul
description: Guides high-level codebase reviews, performance audits under Excel/VBA client-side constraints, architectural modernization, and phased roadmaps.
---

# Codebase Review and Overhaul Guidelines

Use this skill when the user requests an architectural audit, code review, performance optimization, or modernization plan for the VBA codebase.

## Context & Constraints
* **Client-Side execution**: The project is VBA-based and runs entirely on client laptops. 
* **Performance Limitations**: Avoid heavy/fragmented operations that cause Excel to lag, freeze, or exceed memory limits. Prioritize lightweight, optimized execution footprints.

---

## 1. Audit & Review Dimensions

When performing a codebase review, always structure your analysis and suggestions around the following four areas:

### A. Bottleneck & Performance Audit
* **Traps to detect**: Redundant calculation cycles, excessive cell-by-cell sheet reading/writing (encourage array-based transfers), failure to toggle `ScreenUpdating`/`Calculation`/`EnableEvents`, and memory leaks (e.g., objects not set to `Nothing`).
* **Optimizations**: Propose algorithmic improvements to bypass hardware bottlenecks.

### B. Centralized & Unified Architecture
* **Unification**: Detect fragmented, procedural macros and propose refactoring them into the project's interface-driven command pattern (`ICommand`, command handlers).
* **DRY Principle**: Identify duplicate helper routines and push them into infrastructure/library modules (e.g., `Infra_*`, `Lib_*`).

### C. Scalability & Extensibility Framework
* Ensure core architectures (like `AppContainer` and Ribbon bindings) remain decouple-plugged. New tools must be implementable as standalone classes (`FeatCmd_`) without rewriting core features.

### D. Phased Implementation Roadmap
* All proposed restructurings must be broken down into clear, low-risk, and testable phases to avoid breaking existing workbook functionalities.

---

## 2. Proposal Workflow

1. **Structured Suggestions**: Present findings clearly categorized under the four dimensions above.
2. **User Confirmation**: Allow the user to select and approve specific directions.
3. **Implementation Plan**: Create/update the `implementation_plan.md` grouped by component layers.
