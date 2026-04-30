# Beaver Add-in - Gemini Guide

This file is the operational guide for Gemini-style agents working in this repo.
Read `BLUEPRINT.md` first for the project map, then use this file for the work
sequence and decision rules.

Last updated: 2026-04-30

## What This File Is For

Keep this file short, practical, and action-oriented. It should tell an agent
how to work in the repo without re-deriving the entire codebase every time.

Use `BLUEPRINT.md` for stable project facts and architecture.
Use this file for:

- what to inspect first
- how to change files safely
- how to validate changes
- what to avoid touching
- how to report progress

## First Steps

1. Read `BLUEPRINT.md`.
2. Open the target VBA module and its listed `@Dependencies`.
3. Identify the entry path: Ribbon, hotkey, workbook event, or utility.
4. Check whether the change touches `features.json`, `config.json`, or
   `ribbon.xml`.
5. Decide whether the change is source-only or requires `.\Update.ps1`.

## Working Rules

- Edit exported source files, not the compiled workbook.
- Keep `features.json`, `config.json`, `ribbon.xml`, `UI_Ribbon.bas`, and
  `UI_Hotkeys.bas` aligned.
- Treat `FeatCmd_*` classes as the default home for new feature logic.
- Keep Ribbon callbacks and hotkey wrappers thin.
- Use `Infra_Error.Track` at the start of every public macro, callback, and
  workbook event.
- Use `Infra_AppStateGuard` whenever Excel application state changes.
- Use `Infra_Undo.SaveState` before mutating ranges when user undo matters.
- Use `Infra_Progress` for slow work.
- Never touch `Lib_JsonConverter.bas`, `.frx` files, or `_BeaverUndo`/
  `_BeaverTests` if they exist.

## Safe Edit Order

1. Inspect the source module and supporting infra.
2. Make the smallest change that fits the current architecture.
3. Update `features.json` if the UI or command contract changes.
4. Update `config.json` only when a shared setting, icon, or hotkey must change.
5. Keep any new VBA entry point covered by the standard error pattern.
6. Run `.\Update.ps1`.
7. Fix validation, import, compile, or runtime issues.

## Validation Order

When a change is meaningful, validate in this order:

1. Manifest/config consistency.
2. VBA export/import and compilation.
3. Ribbon callback alignment.
4. Runtime smoke tests.

Use `-SkipRuntimeTests` only when you are intentionally deferring execution
checks and you have a clear reason to do so.

## Change Patterns

### New Ribbon command

- Add or update the feature entry in `features.json`.
- Add the matching `Ribbon_On*` callback in `UI_Ribbon.bas`.
- Add or update the matching `CommandName`.
- Add the `FeatCmd_*` implementation.
- Add the resolver case in `AppContainer.ResolveCommand`.
- Add the icon mapping in `config.json` if the control uses a Ribbon icon.

### New hotkey

- Add or update the hotkey entry in `features.json`.
- Add the wrapper in `UI_Hotkeys.bas`.
- Add or update the `CommandName`.
- Add the command implementation or map to an existing command.
- Let `Update.ps1` sync the binding into `config.json`.

### New config value

- Add the key to `config.json`.
- Update `Infra_ConfigModel`.
- Update `Infra_Config`.
- Use the typed accessor everywhere else.

## Output Expectations

- Keep explanations short and concrete.
- Mention exact files changed.
- Call out any validation you ran.
- If something could not be verified, say so plainly.
- If a change affects architecture, mention the follow-up docs that should be
  updated.

## Progress Updates

Provide brief progress updates while working on larger changes. A good update
usually says:

- what has been inspected
- what is being changed next
- whether anything unexpected was found

Keep the tone collaborative and practical.

## Final Check

Before finishing, verify:

- the exported VBA files still match the documented architecture
- generated files were not edited manually
- the ribbon and hotkey contract stayed aligned
- the workbook sync path still succeeds
- the two markdown files stay in sync at the level of intent

## If The Repo Changes

If architecture changes again, update both `BLUEPRINT.md` and this file together.
The blueprint should remain the stable reference; this guide should remain the
operating playbook.
