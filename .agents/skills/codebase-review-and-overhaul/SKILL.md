---
name: Codebase Review and Overhaul
description: Guides high-level codebase reviews, identification of architectural redundancies, suggesting innovative overhauls, checking for bugs/inconsistencies, and making suggestion as a list for the user to review.
---

# Codebase Review and Overhaul Guidelines

Use this skill when the user requests an architectural audit, code review, system modernization, or list of actionable improvement suggestions.

## 1. Audit Dimensions

When reviewing a codebase consider all the aspects including  below mentioned:

### A. Innovative Overhauls & Feature Enhancements
* Think critically about the system's core purpose.
* Propose creative, high-impact improvements (e.g., performance delta-pipelines, caching, UX enhancements).
* Optimising innovation and architectural improvements.
### B. Architecture & Unification
* Look for redundancies (e.g., duplicate helper functions, facade layers that simply forward calls).
* Identify opportunities to merge overlapping scripts or modules.
* Focus on clean dependency injection and modular decoupling.

### C. Bugs, Gaps, & Inconsistencies
* Scan for logical errors, resource leaks, missing cleanups.

---

## 2. Workflow for Proposing Changes

1. **Structured Discovery:** Always list suggestions first and wait for the user to select and approve items.
2. **Implementation Plan:** Draft the `implementation_plan.md` detailing the proposed modifications grouped by component.
